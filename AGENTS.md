# Approved market reporting methodology

The user has permanently approved Sale-to-List % as total Sold Pr divided by
total Orig List Pr for all qualifying closed sales. Use the exact MRED field
`Orig List Pr`. Never substitute regular `List Price`, drop an affected sale,
or average individual percentages. Change this methodology only when the user
explicitly instructs you to do so.

Validate original prices automatically before every report. Flag unusable
values with warnings identifying MLS numbers and property addresses. Withhold
market reports until the values are corrected, and preserve earlier reports.
Display the percentage to two decimal places. Keep README.md and the Excel
Definitions worksheet consistent with this rule. Preserve all other approved
market calculations and community matching rules. Never modify Input exports.

Monthly workflow: Output contains only current reports. Each successful run
creates Archive/YYYY-MM-DD with report files and verified copies of both source
exports. Never overwrite an existing dated archive. Preserve previous Output
contents under Archive/_previous_output. Keep failed validation warnings under
Archive/_validation_warnings, outside current Output. Every JSON includes the
report date. Preserve the double-click Run Reports.cmd workflow.

Archive/market_history.xlsx is append-only by report date. Preserve prior rows,
reject any existing report date and duplicate Report Date + Community keys,
and append only after validation and successful report publication. Keep its
Definitions sheet consistent with approved methodology. Import historical rows
from approved archives without regenerating or altering those monthly reports.

Public community JSON uses only PUBLIC_COMMUNITY_FIELDS in market_report.py.
Never publish MLS identifiers, listing addresses, totals, validation details or
other internal fields. Full records stay in Output/audit/communities and future
private dated archives. Existing historical archives remain unchanged. Only
Output/communities/*.json is intended for public website publication.

Public JSON includes the approved latest-month fields: latest_complete_month (YYYY-MM),
latest_month_closed_sales, latest_month_median_sold_price,
latest_month_sale_to_list_percentage and latest_month_median_closed_dom. Use
the complete calendar month preceding the report date, with the same aggregate
Sold Pr / Orig List Pr methodology. Zero monthly sales require three null
monthly statistics. Preserve all existing fields and history columns.

GitHub publishing remains disabled in publishing.json until the user explicitly
authorizes activation. Never enable it merely to test. Destination is fixed to
colleendahlstrom-png/55-plus-market-data, main, data/, with only the seven named
public JSON files as payloads. Run only after successful monthly/archive/history
completion and public JSON validation. All publishing tests must mock GitHub.
Publication failure must never undo successful local outputs. Use GitHub CLI
browser login and verified secure credential storage; never store credentials
in this project. Setup GitHub.cmd authenticates only and does not activate.

Public JSON now has exactly 23 fields. monthly_history contains exactly 12 chronological, consecutive complete calendar months with only month, closed_sales, median_sold_price, sale_to_original_list_percentage, median_closed_dom. Use aggregate Sold Pr / Orig List Pr * 100, zero/null rules, and reconcile monthly counts and the final month with the existing public metrics. Monthly report dates must be first-of-month. Preserve historical/audit files and all existing public values during the approved October enhancement; use archived October source copies read-only for that enhancement.

The two approved source dates are inventory_updated and closed_sales_through, strict YYYY-MM-DD. closed_sales_through equals closed_period_end (coverage cutoff, not maximum observed sale date). Daily inventory updates use inventory_update.py and change only four inventory metrics plus inventory_updated; preserve all sold fields and closed_sales_through. Monthly runs use the inventory refresh date independently, preserving a known date for an unchanged monthly source or accepting --inventory-date; newly processed sources default to the processing date. Never publish during daily refresh; separate explicit publication is required. Existing archives, source exports, history, and audit records are protected.
