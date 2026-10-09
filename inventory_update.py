"""Refresh current inventory JSON only; never publish or read closed-sales exports."""
import argparse
from datetime import date
import json
import os
from pathlib import Path
import sys
from contextlib import contextmanager
from contextvars import ContextVar

import market_report as engine
import github_publisher as publisher

INVENTORY_FIELDS = ('active_listings', 'median_active_list_price', 'average_active_dom',
                    'pending_contingent_fin_listings', 'inventory_updated')

_BOUNDARY = ContextVar('inventory_write_boundary', default=None)


def _checked_target(path, allowed):
    path = Path(path).absolute()
    if path not in allowed:
        raise engine.ReportError('Daily inventory write outside approved JSON files blocked.')
    for entry in (path.parent.parent, path.parent, path):
        publisher.reject_link(entry)
    if not path.is_file() or path.stat().st_nlink != 1 or path.stat().st_size > 16384 or path.resolve() != path:
        raise engine.ReportError('Daily inventory requires existing, unlinked JSON files.')
    return path


def _audit_boundary(event, args):
    policy = _BOUNDARY.get()
    if policy is None:
        return
    root, readable, writable, written = policy
    if event == 'open':
        name, mode, flags = args
        if not isinstance(name, (str, bytes, os.PathLike)):
            raise engine.ReportError('Descriptor-based file access blocked during daily update.')
        path = Path(os.fsdecode(name)).absolute()
        writing = bool(flags & (os.O_WRONLY | os.O_RDWR | os.O_CREAT | os.O_TRUNC | os.O_APPEND))
        if writing:
            _checked_target(path, writable)
            written.add(path)
        elif path.is_relative_to(root) and path not in readable:
            raise engine.ReportError('Daily inventory access to a protected project file blocked.')
    elif event in ('os.remove', 'os.rename', 'os.mkdir', 'os.rmdir', 'os.chmod', 'os.chown',
                   'os.utime', 'os.link', 'os.symlink', 'os.truncate', 'shutil.copyfile',
                   'subprocess.Popen', 'os.system', 'socket.connect', 'os.chdir'):
        raise engine.ReportError('Filesystem mutation or external action blocked during daily update.')


sys.addaudithook(_audit_boundary)


@contextmanager
def _write_boundary(root, readable, writable):
    written = set()
    token = _BOUNDARY.set((root, readable, writable, written))
    try:
        yield written
    finally:
        _BOUNDARY.reset(token)


def _preserved(original, updated):
    if set(original) != set(updated):
        raise engine.ReportError('Daily update cannot change public field names.')
    for key in original:
        if key not in INVENTORY_FIELDS and json.dumps(original[key], sort_keys=True) != json.dumps(updated[key], sort_keys=True):
            raise engine.ReportError(f'Daily update attempted to change protected field: {key}')


def run_inventory_update(root=engine.BASE, update_date=None):
    root = Path(root).resolve()
    writable = {root / 'Output' / 'communities' / name for name in publisher.COMMUNITIES}
    for path in writable:
        _checked_target(path, writable)
    settings_path = root / 'settings.json'
    settings = json.loads(settings_path.read_text(encoding='utf-8'))
    source = root / 'Input' / settings['input_files']['inventory']
    if source.parent.resolve() != root / 'Input' or source.suffix.lower() != '.xlsx':
        raise engine.ReportError('Inventory source must be an Excel file directly inside Input.')
    publisher.reject_link(source)
    configured = [c['slug'] + '.json' for c in settings['communities']]
    if len(configured) != 7 or set(configured) != set(publisher.COMMUNITIES):
        raise engine.ReportError('Daily inventory requires exactly the seven approved communities.')
    with _write_boundary(root, writable | {settings_path, source}, writable) as written:
        return _run_inventory_update(root, update_date, settings, source, written, writable)


def _run_inventory_update(root, update_date, settings, source, written, writable):
    update_date = update_date or date.today()
    if not settings.get('rules_confirmed'):
        raise engine.ReportError('Reporting rules must be confirmed.')
    output = root / 'Output'
    original = {name: (output / 'communities' / name).read_bytes() for name in publisher.COMMUNITIES}
    payloads = {name: content.decode('utf-8') for name, content in original.items()}
    metadata = {'report_date': publisher.strict_json(payloads['carillon.json'])['report_date']}
    _, old = publisher.validate_public_payloads(payloads, metadata)
    source_hash = engine.file_hash(source)
    rows, _ = engine.load_rows(source, settings['columns'])
    buckets, _ = engine.classify({'inventory': rows}, settings, update_date)
    updated, changes = {}, {}
    for community in settings['communities']:
        name = community['slug'] + '.json'
        row = dict(old[name])
        if update_date < date.fromisoformat(row['inventory_updated']):
            raise engine.ReportError('Inventory update date cannot precede the current inventory refresh date.')
        calculated = engine.summarize(community, buckets[community['slug']],
                                      date.fromisoformat(row['report_date']), update_date)
        row.update({key: calculated[key] for key in INVENTORY_FIELDS})
        _preserved(publisher.strict_json(original[name].decode('utf-8')), row)
        changes[name] = {key: {'before': old[name][key], 'after': row[key]} for key in INVENTORY_FIELDS}
        updated[name] = (json.dumps(row, indent=2, allow_nan=False) + '\n').encode('utf-8')
    publisher.validate_public_payloads({n: b.decode('utf-8') for n, b in updated.items()}, metadata)
    if engine.file_hash(source) != source_hash:
        raise engine.ReportError('Inventory source changed while being read; no JSON updated.')
    if any((output / 'communities' / name).read_bytes() != content for name, content in original.items()):
        raise engine.ReportError('Current JSON changed during inventory update; stopped.')
    try:
        for name, content in updated.items():
            if content != original[name]:
                (output / 'communities' / name).write_bytes(content)
        actual = {name: (output / 'communities' / name).read_bytes() for name in original}
        if actual != updated or not written <= writable:
            raise engine.ReportError('Daily write verification failed.')
        _, verified = publisher.validate_public_payloads({n: b.decode('utf-8') for n, b in actual.items()}, metadata)
        for name, row in verified.items():
            _preserved(publisher.strict_json(original[name].decode('utf-8')), row)
    except Exception:
        for path in written:
            path.write_bytes(original[path.name])
        raise
    return changes


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--inventory-date', type=date.fromisoformat, default=date.today(),
                        help='Date of this inventory update (default: today).')
    args = parser.parse_args()
    try:
        changes = run_inventory_update(update_date=args.inventory_date)
    except (engine.ReportError, publisher.PublicationError, OSError, ValueError, KeyError) as error:
        print(f'Inventory update failed: {error}')
        return 1
    print(json.dumps(changes, indent=2))
    print('Inventory JSON validated. Sold data preserved. Nothing published to GitHub.')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
