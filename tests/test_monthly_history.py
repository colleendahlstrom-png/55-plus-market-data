from copy import deepcopy
from datetime import date
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import market_report as engine
import github_publisher as publisher


class MonthlyHistoryTests(unittest.TestCase):
    def setUp(self):
        self.settings = json.loads((engine.BASE / 'settings.json').read_text())

    def sale(self, key, day, sold=100000, original=200000, dom=10):
        return dict(mls=key, status='CLSD', community='Carillon', city='Plainfield',
                    closed_date=day, sold_price=sold, original_list_price=original,
                    list_price=999999, dom=dom, location='test row', source='test', row=2)

    def report(self, sales=(), as_of=date(2026, 10, 1)):
        buckets, _ = engine.classify({'closed': list(sales)}, self.settings, as_of)
        return engine.summarize(self.settings['communities'][0], buckets['carillon'], as_of)

    def test_exact_months_zero_sales_and_nested_allowlist(self):
        report = self.report()
        history = report['monthly_history']
        self.assertEqual([r['month'] for r in history],
                         ['2025-10', '2025-11', '2025-12'] + [f'2026-{m:02d}' for m in range(1, 10)])
        for row in history:
            self.assertEqual(set(row), {'month', 'closed_sales', 'median_sold_price',
                                      'sale_to_original_list_percentage', 'median_closed_dom'})
            self.assertEqual(row['closed_sales'], 0)
            self.assertTrue(all(row[k] is None for k in engine.MONTHLY_HISTORY_FIELDS[2:]))
        self.assertEqual(len(engine.public_community_report(report)), 23)

    def test_aggregate_not_list_price_or_individual_ratios(self):
        report = self.report([self.sale('a', '2026-09-01'),
                              self.sale('b', '2026-09-30', 300000, 300000, 30)])
        self.assertEqual(report['monthly_history'][-1], dict(month='2026-09', closed_sales=2,
                         median_sold_price=200000, sale_to_original_list_percentage=80, median_closed_dom=20))
        self.assertEqual(sum(r['closed_sales'] for r in report['monthly_history']), report['closed_sales_12_months'])

    def test_existing_twenty_values_unchanged(self):
        result = engine.public_community_report(self.report([self.sale('a', '2026-09-01')]))
        expected = dict(community='Carillon', city='Plainfield', state='Illinois',
                        report_date='2026-10-01', closed_period_start='2025-10-01', closed_period_end='2026-09-30',
                        active_listings=0, median_active_list_price=None, average_active_dom=None,
                        pending_contingent_fin_listings=0, closed_sales_12_months=1, median_sold_price=100000,
                        aggregate_sale_to_list_percentage=50, median_closed_dom=10,
                        sale_to_list_basis='Total Sold Price divided by Total Original List Price for qualifying closed sales.',
                        latest_complete_month='2026-09', latest_month_closed_sales=1,
                        latest_month_median_sold_price=100000, latest_month_sale_to_list_percentage=50,
                        latest_month_median_closed_dom=10)
        self.assertEqual({k: v for k, v in result.items() if k not in ('monthly_history', 'inventory_updated', 'closed_sales_through')}, expected)

    def test_january_year_rollover(self):
        result = self.report([self.sale('a', '2026-12-31'), self.sale('b', '2027-01-01')], date(2027, 1, 1))
        self.assertEqual([r['month'] for r in result['monthly_history']], [f'2026-{m:02d}' for m in range(1, 13)])
        self.assertEqual(result['monthly_history'][-1]['closed_sales'], 1)

    def test_leap_year_and_window_boundaries(self):
        dates = ['2023-02-28', '2023-03-01', '2024-02-29', '2024-03-01']
        result = self.report([self.sale(str(i), d) for i, d in enumerate(dates)], date(2024, 3, 1))
        self.assertEqual(result['monthly_history'][0]['closed_sales'], 1)
        self.assertEqual(result['monthly_history'][-1]['closed_sales'], 1)
        self.assertEqual(result['closed_sales_12_months'], 2)

    def test_annual_medians_are_not_monthly_medians(self):
        sales = [self.sale('a', '2026-08-01', 100000, dom=1),
                 self.sale('b', '2026-08-02', 200000, dom=2),
                 self.sale('c', '2026-08-03', 300000, dom=3),
                 self.sale('d', '2026-09-01', 900000, dom=90)]
        result = self.report(sales)
        self.assertEqual(result['median_sold_price'], 250000)
        self.assertEqual(result['median_closed_dom'], 2.5)
        self.assertEqual(result['monthly_history'][-2]['median_sold_price'], 200000)
        self.assertEqual(result['monthly_history'][-1]['median_sold_price'], 900000)

    def test_invalid_original_prices_in_earlier_month_block_report(self):
        for value in (None, '', 0, -1, 'invalid', 'NaN', 'Infinity', True):
            with self.subTest(value=value), self.assertRaises(engine.OriginalPriceError):
                self.report([self.sale('bad', '2025-10-01', original=value)])

    def test_first_of_month_rejected_before_any_work(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with patch.object(engine, 'load_rows') as load:
                with self.assertRaisesRegex(engine.ReportError, 'first day'):
                    engine.run(root / 'Input', root / 'Output', root / 'settings.json', date(2026, 10, 2))
                load.assert_not_called()
            self.assertEqual(list(root.iterdir()), [])

    def test_invalid_history_rejected_before_github(self):
        valid = self.report([self.sale('a', '2026-09-01')])
        invalid = []
        def case(edit):
            row = deepcopy(valid)
            edit(row)
            invalid.append(row)
        case(lambda r: r.pop('monthly_history'))
        case(lambda r: r.update(monthly_history=None))
        case(lambda r: r['monthly_history'].pop())
        case(lambda r: r['monthly_history'].append(deepcopy(r['monthly_history'][-1])))
        case(lambda r: r['monthly_history'].reverse())
        case(lambda r: r['monthly_history'][1].update(month='2025-10'))
        case(lambda r: r['monthly_history'][1].update(month='2025-12'))
        case(lambda r: r['monthly_history'][1].update(month='2025-1'))
        case(lambda r: r['monthly_history'][0].update(mls='private'))
        case(lambda r: r['monthly_history'][0].pop('median_closed_dom'))
        case(lambda r: r['monthly_history'][0].update(median_sold_price=1))
        for value in (-1, True, 1.5, '1'):
            case(lambda r, v=value: r['monthly_history'][-1].update(closed_sales=v))
        for field in engine.MONTHLY_HISTORY_FIELDS[2:]:
            for value in (None, 'private', True, float('nan'), float('inf'), -1):
                case(lambda r, f=field, v=value: r['monthly_history'][-1].update({f: v}))
        case(lambda r: r['monthly_history'][-1].update(sale_to_original_list_percentage=0))
        case(lambda r: r['monthly_history'][-1].update(median_sold_price=0))
        case(lambda r: r.update(closed_sales_12_months=2))
        case(lambda r: r.update(latest_month_median_closed_dom=20))
        case(lambda r: r.update(latest_complete_month='2026-08'))
        case(lambda r: r.update(closed_period_start='2025-09-01'))
        case(lambda r: r.update(closed_period_end='2026-09-29'))
        case(lambda r: r.update(report_date='2026-10-02'))
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            folder = root / 'Output' / 'communities'
            folder.mkdir(parents=True)
            for name in publisher.COMMUNITIES:
                (folder / name).write_text('{}')
            config = root / 'publishing.json'
            config.write_text('{"enabled": true}')
            for i, report in enumerate(invalid):
                with self.subTest(case=i):
                    public = {k: report[k] for k in engine.PUBLIC_COMMUNITY_FIELDS if k in report}
                    (folder / 'carillon.json').write_text(json.dumps(public))
                    with patch.object(publisher, 'GitHubCLI') as client:
                        result = publisher.after_monthly(root, root / 'Output', {'report_date': '2026-10-01'},
                                                         config, client_factory=client)
                        self.assertEqual(result['json_validation'], 'failed')
                        client.assert_not_called()


if __name__ == '__main__':
    unittest.main()
