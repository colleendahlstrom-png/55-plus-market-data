"""Refresh current inventory JSON only; never publish or read closed-sales exports."""
import argparse
from datetime import date
import json
from pathlib import Path
import tempfile

import market_report as engine
import github_publisher as publisher

INVENTORY_FIELDS = ('active_listings', 'median_active_list_price', 'average_active_dom',
                    'pending_contingent_fin_listings', 'inventory_updated')


def run_inventory_update(root=engine.BASE, update_date=None):
    root = Path(root).resolve()
    update_date = update_date or date.today()
    settings = json.loads((root / 'settings.json').read_text(encoding='utf-8'))
    if not settings.get('rules_confirmed'):
        raise engine.ReportError('Reporting rules must be confirmed.')
    output = root / 'Output'
    metadata = json.loads((output / 'run_info.json').read_text(encoding='utf-8'))
    _, old = publisher.validate_public_json(root, output, metadata)
    source = root / 'Input' / settings['input_files']['inventory']
    if source.parent.resolve() != root / 'Input' or source.suffix.lower() != '.xlsx':
        raise engine.ReportError('Inventory source must be an Excel file directly inside Input.')
    source_hash = engine.file_hash(source)
    rows, _ = engine.load_rows(source, settings['columns'])
    buckets, _ = engine.classify({'inventory': rows}, settings, update_date)
    original = {name: (output / 'communities' / name).read_bytes() for name in old}
    updated, changes = {}, {}
    for community in settings['communities']:
        name = community['slug'] + '.json'
        row = dict(old[name])
        if update_date < date.fromisoformat(row['inventory_updated']):
            raise engine.ReportError('Inventory update date cannot precede the current inventory refresh date.')
        calculated = engine.summarize(community, buckets[community['slug']],
                                      date.fromisoformat(row['report_date']), update_date)
        row.update({key: calculated[key] for key in INVENTORY_FIELDS})
        changes[name] = {key: {'before': old[name][key], 'after': row[key]} for key in INVENTORY_FIELDS}
        updated[name] = (json.dumps(row, indent=2, allow_nan=False) + '\n').encode('utf-8')
    with tempfile.TemporaryDirectory() as temporary:
        stage = Path(temporary)
        folder = stage / 'Output' / 'communities'
        folder.mkdir(parents=True)
        for name, content in updated.items():
            (folder / name).write_bytes(content)
        publisher.validate_public_json(stage, stage / 'Output', metadata)
        if engine.file_hash(source) != source_hash:
            raise engine.ReportError('Inventory source changed while being read; no JSON updated.')
        if any((output / 'communities' / name).read_bytes() != content for name, content in original.items()):
            raise engine.ReportError('Current JSON changed during inventory update; stopped.')
        try:
            for name, content in updated.items():
                if content != original[name]:
                    (output / 'communities' / name).write_bytes(content)
            publisher.validate_public_json(root, output, metadata)
        except Exception:
            for name, content in original.items():
                (output / 'communities' / name).write_bytes(content)
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
