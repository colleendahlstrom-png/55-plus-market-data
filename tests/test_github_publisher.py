from contextlib import redirect_stdout, redirect_stderr
from datetime import date
import io
import json
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import Mock, patch

import openpyxl
import github_publisher as pub
import market_report as engine


class FakeGitHub:
    def __init__(self, failure=None, same=False):
        self.calls = []
        self.failure = failure
        self.same = same

    def authenticate(self):
        self.calls.append(('AUTH', '', None))
        if self.failure == 'AUTH':
            raise pub.PublicationError('Authentication failed.')

    def api(self, method, endpoint, payload=None):
        self.calls.append((method, endpoint, payload))
        if method == self.failure:
            raise pub.PublicationError('Mock GitHub failure.')
        if endpoint.endswith('ref/heads/main'):
            return {'object': {'sha': 'base'}}
        if endpoint.endswith('commits/base'):
            return {'tree': {'sha': 'old-tree'}}
        if endpoint.endswith('trees'):
            return {'sha': 'old-tree' if self.same else 'new-tree'}
        if endpoint.endswith('commits'):
            return {'sha': 'new-commit'}
        if endpoint.endswith('refs/heads/main'):
            return {'object': {'sha': 'new-commit'}}
        raise AssertionError(endpoint)


class PublisherTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.config = self.root / 'publishing.json'
        self.config.write_text('{"enabled": true}')
        self.settings = json.loads((engine.BASE / 'settings.json').read_text())
        settings_path = self.root / 'settings.json'
        settings_path.write_text(json.dumps(self.settings))
        inputs = self.root / 'Input'
        inputs.mkdir()
        for role, filename in self.settings['input_files'].items():
            columns = dict(self.settings['columns'])
            if role == 'closed':
                columns.update(self.settings['closed_columns'])
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(list(columns.values()))
            for i, community in enumerate(self.settings['communities']):
                row = dict(community=community['subdivisions'][0], city=community['city'],
                           mls=role + str(i), status='CLSD' if role == 'closed' else 'ACTV',
                           list_price=300000, original_list_price=320000, sold_price=310000,
                           closed_date=date(2026, 9, 15), dom=20, street_number=123,
                           street_name='Private Address', street_suffix='St', unit=None)
                ws.append([row.get(key) for key in columns])
            wb.save(inputs / filename)
            wb.close()
        self.output, self.meta = engine.run(inputs, self.root / 'Output', settings_path, date(2026, 10, 1))
        self.network_guard = patch('subprocess.run', side_effect=AssertionError('Real subprocess forbidden in publishing tests'))
        self.network_guard.start()
        self.addCleanup(self.network_guard.stop)

    def execute(self, client=None):
        client = client or FakeGitHub()
        return pub.after_monthly(self.root, self.output, self.meta, self.config, lambda: client)

    def snapshot(self):
        return {str(p): engine.file_hash(p) for folder in ('Input', 'Output', 'Archive')
                for p in (self.root / folder).rglob('*') if p.is_file()}

    def test_disabled_makes_no_github_calls(self):
        self.config.write_text('{"enabled": false}')
        factory = Mock(side_effect=AssertionError('Must not initialize authentication'))
        result = pub.after_monthly(self.root, self.output, self.meta, self.config, factory)
        self.assertEqual(result['github_publication'], 'disabled')
        self.assertEqual(result['json_validation'], 'successful')
        factory.assert_not_called()

    def test_one_commit_only_seven_allowlisted_files(self):
        before = self.snapshot()
        client = FakeGitHub()
        result = self.execute(client)
        self.assertEqual(result['github_publication'], 'successful')
        tree = next(payload for method, endpoint, payload in client.calls if endpoint.endswith('/trees'))
        self.assertEqual({entry['path'] for entry in tree['tree']}, {'data/' + f for f in pub.COMMUNITIES})
        for entry in tree['tree']:
            self.assertEqual(set(json.loads(entry['content'])), set(engine.PUBLIC_COMMUNITY_FIELDS))
            self.assertNotIn('Private Address', entry['content'])
        updates = [p for method, _, p in client.calls if method == 'PATCH']
        self.assertEqual(updates, [{'sha': 'new-commit', 'force': False}])
        self.assertEqual(before, self.snapshot())

    def test_failures_preserve_local_success(self):
        for failure in ('AUTH', 'POST', 'PATCH'):
            with self.subTest(failure=failure):
                before = self.snapshot()
                result = self.execute(FakeGitHub(failure=failure))
                self.assertEqual(result['github_publication'], 'failed')
                self.assertEqual(result['history'], 'successful')
                self.assertEqual(before, self.snapshot())

    def test_identical_remote_does_not_create_commit(self):
        client = FakeGitHub(same=True)
        self.assertEqual(self.execute(client)['github_publication'], 'successful')
        self.assertFalse(any(method == 'PATCH' for method, _, _ in client.calls))

    def test_invalid_public_payloads_never_contact_github(self):
        path = self.output / 'communities' / 'carillon.json'
        original = path.read_text()
        cases = [dict(included_mls_numbers=['private']), dict(active_listings=True),
                 dict(median_sold_price='private text'), dict(community='Private address'),
                 dict(report_date='2025-10-01'), dict(median_closed_dom=float('nan'))]
        for change in cases:
            with self.subTest(change=change):
                row = json.loads(original)
                row.update(change)
                path.write_text(json.dumps(row))
                client = FakeGitHub()
                self.assertEqual(self.execute(client)['github_publication'], 'failed')
                self.assertEqual(client.calls, [])
        path.write_text(original)
        path.write_text(original.replace('"city":', '"city":"Lockport", "city":'))
        self.assertEqual(self.execute()['json_validation'], 'failed')

    def test_extra_missing_files_and_wrong_directory_rejected(self):
        folder = self.output / 'communities'
        extra = folder / 'README.md'
        extra.write_text('private')
        client = FakeGitHub()
        self.assertEqual(self.execute(client)['github_publication'], 'failed')
        self.assertEqual(client.calls, [])
        extra.unlink()
        (folder / 'carillon.json').unlink()
        self.assertEqual(self.execute()['json_validation'], 'failed')
        with self.assertRaises(pub.PublicationError):
            pub.validate_public_json(self.root, self.root / 'Archive', self.meta)

    def test_latest_month_public_validation(self):
        path = self.output / 'communities' / 'carillon.json'
        original = json.loads(path.read_text())
        cases = [dict(latest_complete_month='2026-9'), dict(latest_complete_month='2026-08'),
                 dict(latest_month_closed_sales=-1), dict(latest_month_closed_sales=True),
                 dict(latest_month_closed_sales=2), dict(latest_month_closed_sales=0),
                 dict(latest_month_median_sold_price=None), dict(latest_month_median_sold_price='MLS private'),
                 dict(latest_month_sale_to_list_percentage=float('nan')),
                 dict(latest_month_sale_to_list_percentage=0), dict(latest_month_median_closed_dom=-1)]
        for changes in cases:
            with self.subTest(changes=changes):
                path.write_text(json.dumps({**original, **changes}))
                client = FakeGitHub()
                self.assertEqual(self.execute(client)['json_validation'], 'failed')
                self.assertEqual(client.calls, [])
        zero = {**original, 'latest_month_closed_sales': 0, 'latest_month_median_sold_price': None,
                'latest_month_sale_to_list_percentage': None, 'latest_month_median_closed_dom': None}
        zero['monthly_history'] = json.loads(json.dumps(original['monthly_history']))
        zero['monthly_history'][-2].update({k: v for k, v in zero['monthly_history'][-1].items() if k != 'month'})
        zero['monthly_history'][-1].update(closed_sales=0, median_sold_price=None, sale_to_original_list_percentage=None, median_closed_dom=None)
        path.write_text(json.dumps(zero))
        _, validated = pub.validate_public_json(self.root, self.output, self.meta)
        self.assertEqual(validated['carillon.json'], zero)

    def test_incomplete_archive_history_and_warning_gates(self):
        client = FakeGitHub()
        self.meta['warnings'] = ['Review missing data']
        self.assertEqual(self.execute(client)['github_publication'], 'failed')
        self.assertEqual(client.calls, [])

    def test_archive_missing_or_modified_blocks_publication(self):
        path = Path(self.meta['archive_folder']) / 'communities' / 'carillon.json'
        path.write_text('{}')
        client = FakeGitHub()
        self.assertEqual(self.execute(client)['github_publication'], 'failed')
        self.assertEqual(client.calls, [])
        path.unlink()
        self.assertEqual(self.execute(client)['github_publication'], 'failed')
        self.assertEqual(client.calls, [])

    def test_linked_files_are_rejected(self):
        import os
        path = self.output / 'communities' / 'carillon.json'
        linked = self.root / 'linked.json'
        os.link(path, linked)
        client = FakeGitHub()
        self.assertEqual(self.execute(client)['github_publication'], 'failed')
        self.assertEqual(client.calls, [])

    def test_project_configuration_remains_disabled(self):
        self.assertEqual(json.loads((engine.BASE / 'publishing.json').read_text()), {'enabled': False})

    def test_transport_uses_fixed_host_stdin_and_never_echoes_payload(self):
        client = object.__new__(pub.GitHubCLI)
        client.executable, client.env = 'gh.exe', {}
        with patch('subprocess.run', return_value=subprocess.CompletedProcess([], 0, '{"sha":"x"}', '')) as command:
            self.assertEqual(client.api('POST', f'repos/{pub.REPOSITORY}/git/trees', {'tree': []}), {'sha': 'x'})
            args, kwargs = command.call_args
            self.assertIn('github.com', args[0])
            self.assertEqual(args[0][-2:], ['--input', '-'])
            self.assertEqual(json.loads(kwargs['input']), {'tree': []})
            self.assertTrue(kwargs['capture_output'])
        with self.assertRaises(pub.PublicationError):
            client.api('POST', 'repos/other/project/git/trees', {})

    def test_history_mismatch_and_current_only_refresh_block_publication(self):
        client = FakeGitHub()
        self.meta['warnings'] = []
        history = self.root / 'Archive' / 'market_history.xlsx'
        wb = openpyxl.load_workbook(history)
        wb['Market History']['D2'] = 999
        wb.save(history)
        wb.close()
        self.assertEqual(self.execute(client)['github_publication'], 'failed')
        self.assertEqual(client.calls, [])
        self.meta['publication_mode'] = 'current-only'
        self.assertEqual(self.execute(client)['github_publication'], 'failed')
        self.assertEqual(client.calls, [])

    def test_failed_monthly_run_never_calls_publisher(self):
        with patch.object(engine, 'run', side_effect=engine.ReportError('Validation failed')), \
             patch('github_publisher.after_monthly') as after, patch('sys.argv', ['market_report.py']), \
             redirect_stdout(io.StringIO()), redirect_stderr(io.StringIO()):
            self.assertEqual(engine.main(), 1)
            after.assert_not_called()

    def test_main_calls_publisher_after_success_and_failure_is_separate(self):
        result = {'json_validation': 'successful', 'github_publication': 'failed', 'message': 'Mock failure'}
        with patch.object(engine, 'run', return_value=(self.output, self.meta)) as run, \
             patch('github_publisher.after_monthly', return_value=result) as after, \
             patch('sys.argv', ['market_report.py']), redirect_stdout(io.StringIO()) as text:
            self.assertEqual(engine.main(), 2)
            run.assert_called_once()
            after.assert_called_once()
            self.assertIn('Monthly report: SUCCESSFUL', text.getvalue())
            self.assertIn('GitHub publication: FAILED', text.getvalue())

    def test_auth_rejects_plaintext_and_json_state_errors(self):
        client = object.__new__(pub.GitHubCLI)
        for source, state in [('hosts.yml', 'success'), ('keyring', 'error')]:
            client.command = Mock(return_value={'hosts': {'github.com': [dict(tokenSource=source, state=state, active=True)]}})
            with self.assertRaises(pub.PublicationError):
                client.authenticate()
        client.command = Mock(return_value={'hosts': {'github.com': [dict(tokenSource='keyring', state='success', active=True)]}})
        client.authenticate()

    def test_auth_environment_and_error_redaction(self):
        with patch('shutil.which', return_value='gh.exe'), patch.dict('os.environ', {'GH_TOKEN': 'secret'}):
            with self.assertRaises(pub.PublicationError):
                pub.GitHubCLI()
        client = object.__new__(pub.GitHubCLI)
        client.executable, client.env = 'gh.exe', {}
        failed = subprocess.CompletedProcess([], 1, '', 'SECRET HTTP 403 private address')
        with patch('subprocess.run', return_value=failed):
            with self.assertRaises(pub.PublicationError) as raised:
                client.command(['api'])
            self.assertNotIn('SECRET', str(raised.exception))
            self.assertNotIn('private address', str(raised.exception))
        with patch('subprocess.run', side_effect=subprocess.TimeoutExpired('gh', 45)):
            with self.assertRaisesRegex(pub.PublicationError, 'timed out'):
                client.command(['api'])


if __name__ == '__main__':
    unittest.main()
