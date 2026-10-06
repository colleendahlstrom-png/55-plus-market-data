"""Publish only approved public summaries. Live publishing is disabled by default."""
from datetime import date, datetime, timedelta
import hashlib
import json
import math
import os
from pathlib import Path
import re
import shutil
import subprocess

import openpyxl

REPOSITORY = 'colleendahlstrom-png/55-plus-market-data'
BRANCH = 'main'
COMMUNITIES = {
    'carillon.json': ('Carillon', 'Plainfield'),
    'carillon-club.json': ('Carillon Club', 'Naperville'),
    'shorewood-glen.json': ('Shorewood Glen', 'Shorewood'),
    'carillon-lakes.json': ('Carillon Lakes', 'Crest Hill'),
    'grand-haven.json': ('Grand Haven', 'Romeoville'),
    'lincoln-prairie.json': ('Lincoln Prairie', 'Aurora'),
    'lago-vista.json': ('Lago Vista', 'Lockport'),
}


class PublicationError(Exception):
    """Safe message that never includes raw credentials, payloads or API output."""


def strict_json(text):
    def pairs(items):
        result = {}
        for key, value in items:
            if key in result:
                raise PublicationError('Duplicate JSON key; publication blocked.')
            result[key] = value
        return result
    def constant(_):
        raise PublicationError('Nonfinite JSON number; publication blocked.')
    try:
        return json.loads(text, object_pairs_hook=pairs, parse_constant=constant)
    except (ValueError, TypeError):
        raise PublicationError('Invalid JSON; publication blocked.') from None


def reject_link(path):
    if path.is_symlink() or (hasattr(path, 'is_junction') and path.is_junction()):
        raise PublicationError('Linked publication paths are not permitted.')


def validate_public_json(project_root, output, metadata):
    import market_report as engine
    root = Path(project_root).resolve()
    output = Path(output).absolute()
    reject_link(output)
    if output.resolve() != root / 'Output':
        raise PublicationError('Publication is restricted to this project’s Output/communities.')
    folder = output / 'communities'
    reject_link(folder)
    if set(p.name for p in folder.iterdir()) != set(COMMUNITIES):
        raise PublicationError('Expected exactly the seven approved public JSON files.')
    day = date.fromisoformat(metadata['report_date'])
    payloads, reports = {}, {}
    counts = ('active_listings', 'pending_contingent_fin_listings', 'closed_sales_12_months', 'latest_month_closed_sales')
    for name, (community, city) in COMMUNITIES.items():
        path = folder / name
        reject_link(path)
        if not path.is_file() or path.stat().st_nlink != 1 or path.stat().st_size > 16384:
            raise PublicationError('Public files must be small regular files without hard links.')
        text = path.read_text(encoding='utf-8')
        row = strict_json(text)
        if not isinstance(row, dict) or set(row) != set(engine.PUBLIC_COMMUNITY_FIELDS):
            raise PublicationError(f'{name}: public JSON must have exactly the approved 23 fields.')
        try:
            engine.validate_monthly_history(row)
            engine.validate_source_dates(row)
        except engine.ReportError as error:
            raise PublicationError(f'{name}: {error}') from None
        expected_text = dict(community=community, city=city, state='Illinois',
                             report_date=day.isoformat(), closed_period_start=engine.year_before(day).isoformat(),
                             closed_period_end=(day - timedelta(days=1)).isoformat(),
                             sale_to_list_basis=engine.SALE_TO_LIST_DEFINITION,
                             latest_complete_month=engine.latest_month_bounds(day)[0].strftime('%Y-%m'))
        if any(row[k] != value for k, value in expected_text.items()):
            raise PublicationError(f'{name}: identity, period or methodology does not match the approved report.')
        for key in counts:
            if type(row[key]) is not int or row[key] < 0:
                raise PublicationError(f'{name}: invalid count.')
        if row['latest_month_closed_sales'] > row['closed_sales_12_months']:
            raise PublicationError(f'{name}: monthly sales cannot exceed the qualifying 12-month count.')
        for count, metrics in (
            ('active_listings', ('median_active_list_price', 'average_active_dom')),
            ('closed_sales_12_months', ('median_sold_price', 'aggregate_sale_to_list_percentage', 'median_closed_dom')),
            ('latest_month_closed_sales', ('latest_month_median_sold_price', 'latest_month_sale_to_list_percentage', 'latest_month_median_closed_dom')),
        ):
            for key in metrics:
                value = row[key]
                if row[count] == 0:
                    if value is not None:
                        raise PublicationError(f'{name}: unavailable metric must be null.')
                elif type(value) not in (int, float) or not math.isfinite(value) or value < 0:
                    raise PublicationError(f'{name}: invalid market metric.')
                elif ('price' in key or 'percentage' in key) and value == 0:
                    raise PublicationError(f'{name}: invalid price or percentage.')
        payloads[name] = text
        reports[name] = row
    return payloads, reports


def validate_completion(project_root, output, metadata, payloads, reports):
    """Local-only checks. No archive, audit or history content is sent to GitHub."""
    import market_report as engine
    if metadata.get('publication_mode') or metadata.get('warnings'):
        raise PublicationError('Only a completed monthly run without validation warnings may publish.')
    audit = metadata.get('original_list_price_validation', {})
    if (audit.get('field') != 'Orig List Pr' or audit.get('unusable_records') != []
            or audit.get('fallback_used') is not False):
        raise PublicationError('Original-price validation was not successful.')
    archive = Path(metadata['archive_folder'])
    if not archive.is_dir() or archive.name != metadata['report_date']:
        raise PublicationError('Successful dated archive is required.')
    archived = strict_json((archive / 'run_info.json').read_text(encoding='utf-8'))
    current = strict_json((Path(output) / 'run_info.json').read_text(encoding='utf-8'))
    if archived != metadata or current != metadata:
        raise PublicationError('Current report and archive do not match this completed run.')
    if audit.get('qualifying_closed_sales') != sum(r['closed_sales_12_months'] for r in reports.values()):
        raise PublicationError('Closed-sale validation count mismatch.')
    for name, text in payloads.items():
        if (archive / 'communities' / name).read_text(encoding='utf-8') != text:
            raise PublicationError('Public JSON differs from the successfully archived report.')
    wb = openpyxl.load_workbook(archive.parent / 'market_history.xlsx', read_only=True, data_only=True)
    try:
        found = {}
        for row in wb['Market History'].iter_rows(min_row=2, values_only=True):
            if engine.closing_date(row[0], 'history').isoformat() != metadata['report_date']:
                continue
            if row[1] in found:
                raise PublicationError('Duplicate history record; publication blocked.')
            found[row[1]] = row
        if set(found) != {r['community'] for r in reports.values()}:
            raise PublicationError('History does not contain all seven communities for this run.')
        for report in reports.values():
            expected = [report[field] for _, field in engine.HISTORY_COLUMNS][1:]
            actual = list(found[report['community']][1:])
            if actual[8] is not None:
                actual[8] = round(actual[8] * 100, 4)
            if actual != expected:
                raise PublicationError('History values differ from this monthly report.')
    finally:
        wb.close()


class GitHubCLI:
    def __init__(self):
        self.executable = shutil.which('gh')
        if not self.executable:
            raise PublicationError('GitHub CLI is not installed or is not on PATH.')
        if any(os.environ.get(key) for key in ('GH_TOKEN', 'GITHUB_TOKEN', 'GH_ENTERPRISE_TOKEN', 'GITHUB_ENTERPRISE_TOKEN')):
            raise PublicationError('Environment-token overrides are not permitted; use GitHub CLI secure browser login.')
        self.env = os.environ.copy()
        self.env.pop('GH_DEBUG', None)
        self.env.update(GH_HOST='github.com', GH_PROMPT_DISABLED='1', GH_PAGER='cat')

    def command(self, arguments, payload=None):
        try:
            result = subprocess.run([self.executable, *arguments],
                input=json.dumps(payload, allow_nan=False) if payload is not None else None,
                capture_output=True, text=True, encoding='utf-8', timeout=45, env=self.env,
                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        except (OSError, subprocess.TimeoutExpired):
            raise PublicationError('GitHub command failed or timed out; check connection and authentication. If the final update timed out, verify GitHub before retrying.') from None
        if result.returncode:
            status = re.search(r'HTTP (\d{3})', result.stderr or '')
            suffix = f' (HTTP {status.group(1)})' if status else ''
            raise PublicationError('GitHub request failed' + suffix + '; check permissions, branch protection or remote conflicts.')
        return strict_json(result.stdout)

    def authenticate(self):
        status = self.command(['auth', 'status', '--hostname', 'github.com', '--active', '--json', 'hosts'])
        entries = status.get('hosts', {}).get('github.com', [])
        if len(entries) != 1 or entries[0].get('state') != 'success' or entries[0].get('active') is not True:
            raise PublicationError('Sign in to GitHub CLI using browser authentication.')
        if entries[0].get('tokenSource') != 'keyring':
            raise PublicationError('Secure credential-store authentication is required; plaintext fallback is not accepted.')

    def api(self, method, endpoint, payload=None):
        if not endpoint.startswith(f'repos/{REPOSITORY}/'):
            raise PublicationError('Unapproved GitHub destination.')
        args = ['api', '--hostname', 'github.com', '--method', method, endpoint]
        if payload is not None:
            args += ['--input', '-']
        return self.command(args, payload)


def publish_commit(client, payloads, report_date):
    """One tree/commit/reference update; never stage or push the workspace."""
    if set(payloads) != set(COMMUNITIES):
        raise PublicationError('Only the seven approved filenames may publish.')
    client.authenticate()
    prefix = f'repos/{REPOSITORY}/git/'
    base = client.api('GET', prefix + 'ref/heads/main')['object']['sha']
    base_tree = client.api('GET', prefix + 'commits/' + base)['tree']['sha']
    tree = client.api('POST', prefix + 'trees', {
        'base_tree': base_tree,
        'tree': [{'path': 'data/' + name, 'mode': '100644', 'type': 'blob', 'content': payloads[name]}
                 for name in COMMUNITIES],
    })['sha']
    if tree == base_tree:
        return {'status': 'successful', 'message': 'GitHub already contains identical public JSON.', 'commit': base}
    commit = client.api('POST', prefix + 'commits', {
        'message': f'Monthly market report {report_date}', 'tree': tree, 'parents': [base],
    })['sha']
    response = client.api('PATCH', prefix + 'refs/heads/main', {'sha': commit, 'force': False})
    if response.get('object', {}).get('sha') != commit:
        raise PublicationError('GitHub did not confirm the expected commit; verify remote state before retrying.')
    return {'status': 'successful', 'message': 'Seven public JSON files published in one commit.', 'commit': commit}


def after_monthly(project_root, output, metadata, config_path=None, client_factory=GitHubCLI):
    """Called only after run() returns successfully. Never alters local reports/history."""
    result = {'report_date': metadata['report_date'], 'monthly_report': 'successful',
              'archive': 'successful', 'history': 'successful', 'json_validation': 'not checked',
              'github_publication': 'disabled', 'message': 'Live publishing disabled; no GitHub requests made.'}
    try:
        payloads, reports = validate_public_json(project_root, output, metadata)
        result['json_validation'] = 'successful'
        config = strict_json((config_path or Path(project_root) / 'publishing.json').read_text(encoding='utf-8'))
        if set(config) != {'enabled'} or type(config['enabled']) is not bool:
            raise PublicationError('Invalid publishing configuration; publication blocked.')
        if config['enabled']:
            validate_completion(project_root, output, metadata, payloads, reports)
            publication = publish_commit(client_factory(), payloads, metadata['report_date'])
            result.update(github_publication=publication['status'], message=publication['message'])
            result['commit'] = publication['commit']
    except Exception as error:
        # Do not expose raw API responses, JSON content, auth errors or credentials.
        result['github_publication'] = 'failed'
        result['message'] = str(error) if isinstance(error, PublicationError) else 'Publication checks or GitHub access failed; local monthly reports remain saved.'
        if result['json_validation'] == 'not checked':
            result['json_validation'] = 'failed'
    logs = Path(project_root) / 'Logs'
    try:
        logs.mkdir(exist_ok=True)
        path = logs / f"publication_{metadata['report_date']}_{datetime.now():%Y%m%d_%H%M%S_%f}.json"
        path.write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
    except OSError:
        result['log_warning'] = 'Could not save publication log; local monthly reports remain saved.'
    return result
