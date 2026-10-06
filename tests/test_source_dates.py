from datetime import date
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import market_report as engine
import github_publisher as publisher
import inventory_update


class SourceDateTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        from test_market_report import ReportingTests
        helper = ReportingTests()
        helper.setUp()
        self.settings = helper.settings
        self.inputs = helper.make_exports(self.root)
        self.config = self.root / 'settings.json'
        self.config.write_text(json.dumps(self.settings))
        self.output, self.metadata = engine.run(self.inputs, self.root / 'Output', self.config,
                                               date(2026, 10, 1), inventory_date=date(2026, 10, 1))

    def public(self):
        return {p.name: json.loads(p.read_text()) for p in (self.output / 'communities').glob('*.json')}

    def test_daily_updates_date_preserves_all_other_fields_and_protected_files(self):
        before = self.public()
        files = [p for folder in ('Input', 'Archive', 'Output') for p in (self.root / folder).rglob('*')
                 if p.is_file() and p.parent != self.output / 'communities']
        hashes = {p: engine.file_hash(p) for p in files}
        real_load = engine.load_rows
        def inventory_only(path, columns):
            self.assertEqual(path.name, self.settings['input_files']['inventory'])
            return real_load(path, columns)
        with patch.object(engine, 'load_rows', side_effect=inventory_only), \
             patch.object(publisher, 'GitHubCLI', side_effect=AssertionError('No GitHub access')):
            inventory_update.run_inventory_update(self.root, date(2026, 10, 6))
        for name, row in self.public().items():
            self.assertEqual(len(row), 23)
            self.assertEqual(row['inventory_updated'], '2026-10-06')
            self.assertEqual(row['closed_sales_through'], '2026-09-30')
            self.assertEqual({k: v for k, v in row.items() if k != 'inventory_updated'},
                             {k: v for k, v in before[name].items() if k != 'inventory_updated'})
        self.assertEqual(hashes, {p: engine.file_hash(p) for p in files})

    def test_monthly_reused_inventory_preserves_refresh_date(self):
        engine.run(self.inputs, self.output, self.config, date(2026, 11, 1))
        for row in self.public().values():
            self.assertEqual(row['inventory_updated'], '2026-10-01')
            self.assertEqual(row['closed_sales_through'], '2026-10-31')

    def test_monthly_explicit_inventory_snapshot_is_independent_of_report_date(self):
        engine.run(self.inputs, self.output, self.config, date(2026, 11, 1), inventory_date=date(2026, 10, 30))
        for row in self.public().values():
            self.assertEqual(row['inventory_updated'], '2026-10-30')
            self.assertEqual(row['closed_sales_through'], '2026-10-31')

    def test_monthly_new_inventory_defaults_to_processing_date(self):
        # Harmless XLSX ZIP trailing bytes change the source fingerprint.
        path = self.inputs / self.settings['input_files']['inventory']
        path.write_bytes(path.read_bytes() + b'new snapshot')
        engine.run(self.inputs, self.output, self.config, date(2026, 11, 1))
        for row in self.public().values():
            self.assertEqual(row['inventory_updated'], date.today().isoformat())
            self.assertEqual(row['closed_sales_through'], '2026-10-31')

    def test_bad_dates_block_publication_before_network(self):
        path = self.output / 'communities' / 'carillon.json'
        original = path.read_text()
        for field in ('inventory_updated', 'closed_sales_through'):
            for value in (None, '', '20261006', '2026-10-6', '2026-02-30', '2026-10-06T00:00:00', True, 123):
                with self.subTest(field=field, value=value):
                    row = json.loads(original)
                    row[field] = value
                    path.write_text(json.dumps(row))
                    with self.assertRaises(publisher.PublicationError):
                        publisher.validate_public_json(self.root, self.output, self.metadata)
        row = json.loads(original)
        row['closed_sales_through'] = '2026-09-29'
        path.write_text(json.dumps(row))
        with self.assertRaises(publisher.PublicationError):
            publisher.validate_public_json(self.root, self.output, self.metadata)
        for field in ('inventory_updated', 'closed_sales_through'):
            row = json.loads(original)
            del row[field]
            path.write_text(json.dumps(row))
            with self.assertRaises(publisher.PublicationError):
                publisher.validate_public_json(self.root, self.output, self.metadata)

    def test_daily_older_date_rejected_without_changes(self):
        before = self.public()
        with self.assertRaisesRegex(engine.ReportError, 'cannot precede'):
            inventory_update.run_inventory_update(self.root, date(2026, 9, 30))
        self.assertEqual(self.public(), before)


if __name__ == '__main__':
    unittest.main()
