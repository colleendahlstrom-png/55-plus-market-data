# 55+ Community Market Reports

This system reads your two MRED Excel exports and creates a combined Excel summary and a separate JSON report for each of seven communities. **It never changes the original exports.** Every successful run publishes the current reports in `Output` and saves a dated archive with copies of the exact source exports used.

## Each time you receive new exports

1. Check that the previous successful report is saved in `Archive/YYYY-MM-DD`. Its source exports are already copied there automatically. You can then replace the two files in `Input` with the next month's exports.
2. Put the two new `.xlsx` exports in `Input`, using these exact names:
   - `MRED MLS Active Pending Data.xlsx`
   - `MRED Closed Data - Past 12 Mths.xlsx`
3. Include the columns listed below, including **Orig List Pr** in the Closed Sales export, and make sure the MLS searches cover all seven communities. The closed export must cover the full reporting period. If using OneDrive, make sure both files are downloaded locally.
4. Close both exports in Excel. Double-click **Run Reports.cmd** in the main folder.
5. Enter the report date as `YYYY-MM-01` (the first day of the month). Press Enter only when today is the first of the month. Other dates are rejected before files are written.
6. Open **Output/market_summary.xlsx**. The completion message shows the report date, communities processed, qualifying closed sales analyzed, archive location and whether validation warnings occurred. Check `listing_review.csv` for unexpected exclusions.

Example: report date **2026-10-01** includes closings from **2025-10-01 through 2026-09-30**, inclusive. It excludes closings on the report date. This is a rolling calendar-year window, not always 365 days. For February 29, the previous-year boundary is February 28.

Active and pending counts reflect the snapshot in your inventory export. Changing the report date cannot reconstruct past inventory. Use an inventory export collected for the date you want to report. The program cannot prove that your MLS search included every listing or all 12 months, so check your export search criteria.

## What you receive

```text
Output/
  market_summary.xlsx
  communities/                 # Seven current community JSON files
  audit/communities/           # Full private reports, including MLS IDs and audit totals
  listing_review.csv
  run_info.json
Archive/
  market_history.xlsx         # One row per community and report date
  2026-10-01/
    market_summary.xlsx
    communities/
      carillon.json
      carillon-club.json
      shorewood-glen.json
      carillon-lakes.json
      grand-haven.json
      lincoln-prairie.json
    listing_review.csv
    run_info.json
    source_exports/
      MRED MLS Active Pending Data.xlsx
      MRED Closed Data - Past 12 Mths.xlsx
  _previous_output/            # Preserved prior Output contents
  _validation_warnings/        # Warning files for unsuccessful validation
```

The Excel file has one summary table with all seven communities and all eight requested metrics, plus a Definitions worksheet. It is a finished snapshot. Editing it does not update the JSON files; rerun the program with new exports to update the results.

Only **Output/communities/*.json** is intended for public GitHub/website publication. Each of the seven files contains exactly these 23 fields: `community`, `city`, `state`, `report_date`, `closed_period_start`, `closed_period_end`, `active_listings`, `median_active_list_price`, `average_active_dom`, `pending_contingent_fin_listings`, `closed_sales_12_months`, `median_sold_price`, `aggregate_sale_to_list_percentage`, `median_closed_dom`, `sale_to_list_basis`, `latest_complete_month`, `latest_month_closed_sales`, `latest_month_median_sold_price`, `latest_month_sale_to_list_percentage`, `latest_month_median_closed_dom`, `monthly_history`, `inventory_updated`, and `closed_sales_through`.

The five latest-month fields describe the complete calendar month immediately before the report date. For October 1, 2026, `latest_complete_month` is `2026-09`, covering September 1 through September 30. January reports use December of the previous year; February follows the actual calendar, including leap day where applicable.

Only qualifying CLSD sales within that month feed its closed count, median Sold Pr, aggregate Sold Pr divided by Orig List Pr, and median MT. The original-price validation and no-substitution rule apply. The percentage is stored in percentage points with four decimal places, consistent with the existing aggregate percentage. If there are no sales, the monthly count is 0 and the three monthly statistics are null. No previous month is carried forward. All previous public fields and the trailing-12-month calculations keep their existing meaning and values.

Monthly additions are currently in the public and private community JSON. They do not add columns to market_history.xlsx or change existing Excel metrics. The October archive and historical baseline remain unchanged. Future normal monthly runs include the 23-field JSON automatically. Publishing validation requires all 23 approved fields, the correct YYYY-MM month, valid numeric values and the zero/null rules; private fields remain prohibited.

Public JSON excludes MLS identifiers, addresses, individual listing information, audit totals and validation details. An explicit field allowlist prevents new internal fields from appearing publicly. The full community records remain locally in **Output/audit/communities**, alongside the existing private listing review and run metadata. Future dated archives contain both public JSON and private audit copies. Existing historical archives are unchanged and may contain listing identifiers; they are not public publishing folders.

Publish only the seven files inside `Output/communities`, not the entire Output or Archive folder. The project `.gitignore` excludes local source exports, archives and private outputs from ordinary Git staging; it does not remove anything that was already tracked or published.

Percentages are percentage points: `98.5` means **98.5%**. `null` means no qualifying observations. In Excel, that appears as `n.a.`; counts remain zero.

Every JSON file, including `run_info.json`, includes `report_date` in `YYYY-MM-DD` format. Archive folders use that report date, not the date you happen to run the program. Source copies are verified against the recorded SHA-256 fingerprints before publication.

**Existing dated archives are never overwritten.** If you run the same report date again, the program warns you and stops, preserving both the archive and current Output. It does not offer an automatic overwrite option. Review the archive before deciding how to handle a corrected report; do not choose an inaccurate date just to bypass the protection.

`Output` contains only the newest successfully generated reports. Before replacing it, the system preserves its previous contents under `Archive/_previous_output` with a unique timestamp. This also preserves older reports created before dated archives were introduced. Those older reports are not retroactively assigned copies of today's source files. Failed validation warnings are kept outside Output, in `Archive/_validation_warnings`. A publication failure restores the previous Output when possible; no success message is shown unless both the current report and dated archive have been published.

`listing_review.csv` identifies every source row, its community match, and whether it was included or excluded. Other communities in your exports are expected exclusions. Review this file to catch new subdivision spellings or unexpected city names. `run_info.json` records the source names, worksheets, row counts, SHA-256 fingerprints, settings, reporting period and warnings for that run.

## Your historical market database

Open **Archive/market_history.xlsx** to see all completed report dates in one table. The `Market History` sheet contains one row per community per report date, with the same eight approved market metrics as the monthly reports. Use the table's filters to select a community or date. The second sheet, `Definitions`, explains the approved metrics. Dates and numbers are stored as real Excel values; Sale-to-List % displays two decimal places. Blank price, DOM or percentage cells mean there were no qualifying observations, not a value of zero.

Lago Vista is the seventh community, matched by subdivision `Lago Vista` and city `Lockport`, ignoring capitalization and extra spaces. It uses the same eight calculations, including total Sold Pr divided by total Orig List Pr. Future monthly runs create `communities/lago-vista.json` and include Lago Vista in the summary, archive and history.

The October 1 current Output was refreshed to include seven communities at your request. This current-only update is not an archived monthly run: the protected October archive remains unchanged. You subsequently authorized adding exactly one Lago Vista row to market_history.xlsx, so the historical baseline now has seven records while preserving the original six. Normal double-click runs still reject an existing archive/history date.

The initial six rows were imported from the approved October 1, 2026 archive, without rerunning that month's reports. Future successful runs of **Run Reports.cmd** automatically add one row per configured community after data validation. You do not need to copy or paste anything.

**Existing historical rows are never edited by the reporting program.** It checks Report Date + Community keys and rejects duplicate incoming communities. If any row already has the incoming report date, the entire run stops with a warning—even if that date has only some communities. It does not fill gaps, replace rows or create duplicates. Review the history before attempting a correction. Do not manually edit this database as part of routine monthly reporting.

Close `market_history.xlsx` in Excel before running the next monthly report. The system stages the updated history, current reports and archive before publishing them. If history cannot be saved, it stops and restores the prior current reports rather than reporting success. Validation failures do not append history. Sale-to-List % continues to use **total Sold Pr divided by total Orig List Pr**, with no regular List Price fallback. No existing calculation or community rule is changed by storing history.

## Columns inspected in your actual files

The inventory sheet was `ConnectMLS_export (57)` with 114 records. The replacement closed sheet is `ConnectMLS_export (62)` with 322 records and an additional original price field. Its exact header is **Orig List Pr**, rather than the abbreviated “Orig List.” Future worksheet names may change; the program locates the header within the first 20 rows of a single qualifying sheet.

| Meaning | Actual MRED header |
| --- | --- |
| Subdivision/community | `Subdivision` |
| Listing identifier | `MLS #` |
| Status | `Stat` |
| Days on market | `MT` (confirmed by you) |
| Regular list price (used for active price metrics) | `List Price` |
| Original list price (closed sales percentage denominator) | `Orig List Pr` (required only in the closed export) |
| Sold price | `Sold Pr` |
| Closing date | `Closed Date` |
| Street address | `Street #`, `Str Name`, `Sfx`, `Unit #` |
| City | `City` |

The program uses **Orig List Pr** as the original listing price for every qualifying closed sale. It ignores the supplied `SP:LP` column and recomputes the aggregate from actual dollar amounts. **Regular List Price is never used as the percentage denominator or as a fallback.** Address fields are identified and required in the export layout, but are not used to guess community membership. Blank unit numbers are fine.

Before calculating any results, the program checks every qualifying closed sale's original price. Blank, zero, negative, nonnumeric and nonfinite values stop the run. An on-screen warning and a dated `WARNING` text report in `Archive/_validation_warnings` list the affected MLS numbers, property addresses, communities, source rows and unusable values. Market reports are withheld until the problems are corrected. No replacement price is silently chosen, no affected sale is silently dropped, and previous reports remain unchanged. A successful run records the number of validated sales and an empty unusable-record list in `run_info.json`.

**This methodology is permanent and may be changed only on your explicit instruction.** The program rejects a settings change that maps original price to any field other than `Orig List Pr`. Every future run validates this field automatically. Sale-to-List % is displayed to **two decimal places** in Excel; JSON retains four decimal places for audit comparisons.

## Confirmed reporting rules

These rules were confirmed before the initial reports were generated:

- Active: `ACTV`, `NEW`, `PCHG`.
- Pending / contingent / FIN combined: `A/I`, `PEND`, `HC24`, `HC48`, `HS48`, `FIN`, `SS`.
- Closed: `CLSD`, with a closing date inside the reporting window.
- DOM: use `MT`, including valid zero-day observations.
- Match subdivision **and** city exactly after ignoring capitalization and extra spaces. No partial-name or fuzzy matching is used.

| Report community | Subdivision in export | Required city |
| --- | --- | --- |
| Carillon | Carillon | Plainfield |
| Carillon Club | Carillon Club | Naperville |
| Shorewood Glen | Shorewood Glen Del Webb | Shorewood |
| Carillon Lakes | Carillon Lakes | Crest Hill |
| Grand Haven | Grand Haven | Romeoville |
| Lincoln Prairie | Lincoln Prairie | Aurora |
| Lago Vista | Lago Vista | Lockport |

This keeps Carillon North, Carillon at Stonegate and other similarly named developments out of Carillon's report. The exports do not contain a state column; Illinois is the configured report context, with membership determined by subdivision and city.

## How the numbers are calculated

| Metric | Calculation |
| --- | --- |
| Active listings | Count of matched listings in active statuses |
| Median active list price | Middle List Price among active listings; average of the two middle values for an even count |
| Average active DOM | Arithmetic average of active MT values |
| Pending / contingent / FIN | Combined count of matched listings in all confirmed pending statuses |
| Closed sales, previous 12 months | Count of matched CLSD listings inside the date window |
| Median sold price | Median Sold Pr for those closed sales |
| Actual aggregate sale-to-list percentage | Total Sold Price divided by Total Original List Price for qualifying closed sales. |
| Median closed DOM | Median MT for those closed sales |

**Sale-to-List %:** "Total Sold Price divided by Total Original List Price for qualifying closed sales."

The calculation is **100 × sum(Sold Pr) ÷ sum(Orig List Pr)**, expressed as a percentage. It is neither an average nor a median of listing-level percentages. For example, sales of $100,000 and $300,000 against original list prices of $200,000 and $300,000 produce $400,000 ÷ $500,000 = **80%**, not 75%.

The public JSON key `aggregate_sale_to_list_percentage` uses original list prices, and `sale_to_list_basis` identifies the definition. The private audit JSON additionally includes `sale_to_list_denominator_field` and `total_closed_original_list_price` for verification. Its `total_closed_list_price` is an audit total only; it does not feed the percentage. None of those internal fields is included in public JSON.

Calculations use unrounded source dollar amounts. JSON dollar/DOM results are rounded to two decimal places and the percentage to four. Excel displays whole dollars, one decimal for DOM, and two decimals for percentages.

## If the program stops or reports a warning

- **Access denied:** close the export in Excel and check OneDrive download/sync status. Retry.
- **Archive already exists:** that report date is protected. The run stops without overwriting either the archive or current reports. Review the existing archive before taking further action.
- **Wrong file names or extra files:** keep only the two expected `.xlsx` exports in Input. Excel's temporary `~$` lock files are ignored.
- **Missing or changed columns:** export the required fields again. The program stops rather than choosing a different field silently.
- **Unexpected status:** confirm its meaning before changing the status lists in `settings.json`. This check applies to all source rows, even other communities.
- **Duplicate MLS number:** correct duplicate records or overlapping exports. The program stops rather than deciding which row to keep. Separate MLS numbers for the same address remain separate listings.
- **Missing subdivision or city:** correct the export before rerunning. No address-based guess is made.
- **Unusable Orig List Pr:** review every listed MLS number and source row, correct the closed export, and rerun. No reports are generated until every qualifying closed sale has a usable original price. Regular List Price is never substituted.
- **Invalid price, DOM or closing date:** the program identifies the source row. Required metric data for included listings must be valid; prices must be positive and DOM must be nonnegative. Missing values are not replaced with zeros or silently dropped.
- **No matching records warning:** the report shows zero counts and unavailable price/DOM metrics. Verify that your search covered that community; zero records alone cannot establish zero market activity.
- **An unfamiliar subdivision/city appears in listing_review.csv:** decide whether it belongs to a requested community before adding an alias. Unmatched rows are excluded and disclosed; they are never guessed into a community.

A failed run does not publish a new current report. Look for the completion message and verify the report date in Output before using the results.

## One-time setup on another Windows computer

This computer already has the Python runtime and `openpyxl` needed for the launcher. The launcher first looks for the bundled Codex Python, then a standard Python installation.

On another computer, install Python 3.10 or newer with its Windows launcher. Open a terminal in this folder and run:

```powershell
py -3 -m pip install -r requirements.txt
```

Then double-click `Run Reports.cmd`. No Excel installation is required to generate the files. Excel or another spreadsheet viewer is needed to view the `.xlsx` summary.

## GitHub publication (implemented, currently disabled)

Live publishing is **disabled** in `publishing.json` (`enabled` is `false`). Keep it disabled until you explicitly authorize activation. Normal double-click monthly reporting still works without GitHub installed or signed in. No GitHub requests are made while disabled. The existing October current-only refresh is not eligible for automatic monthly publication.

When activated later, only these seven files from `Output/communities` can be sent to `colleendahlstrom-png/55-plus-market-data`, branch `main`, under `data/`: carillon.json, carillon-club.json, shorewood-glen.json, carillon-lakes.json, grand-haven.json, lincoln-prairie.json and lago-vista.json. The destination and filenames are fixed in the publisher. It never pushes the local project repository or uploads directories.

The publisher runs only after the monthly report, dated archive and history update succeed. It validates the exact public field list, types, finite numeric values, community identities and reporting dates, rejects links and unexpected/missing files, and checks that the seven public summaries match the committed archive and history. Data validation warnings block live publication. These are publication checks; approved market calculations and local reporting rules are unchanged.

Publication updates all seven files together in one commit using the existing remote tree, preserving every other repository file. The branch update is not forced. If GitHub reports a conflict, permission problem, expired authentication or network failure, the local report, archive and history remain saved. Identical remote files are reported as already up to date. If a final update times out, verify the repository before retrying because GitHub may have accepted the commit. Do not rerun the same monthly date to retry: duplicate-date protection remains in place. There is no automatic retry or standalone publishing command in this version.

Completion messages separately show Monthly report, Archive, History, JSON validation and GitHub publication status. A GitHub failure returns exit code 2 while explicitly identifying the successful local stages. Private publication logs are written to `Logs` with dates, statuses and a commit identifier when available. Credentials, JSON payloads and raw API responses are not logged. Logs are excluded from ordinary Git staging.

### One-time authentication setup, when ready

1. Install the current GitHub CLI for Windows from [cli.github.com](https://cli.github.com/). Reopen your terminal afterward.
2. Double-click **Setup GitHub.cmd**. Complete the browser sign-in using an account with write access to this repository.
3. Setup verifies that GitHub CLI reports secure system credential storage (`keyring`). It rejects plaintext storage and token environment overrides. No password or access token is stored in project files. If secure storage cannot be verified, do not activate publication.
4. Authentication does not activate publishing. A separate explicit authorization is required before changing `publishing.json` to enable live monthly publishing.

The setup uses `gh auth login --hostname github.com --git-protocol https --web`; see the [GitHub CLI authentication documentation](https://cli.github.com/manual/gh_auth_login). Neither setup nor tests publish reports. Tests use mocked GitHub responses and temporary local report fixtures. Actual repository access and branch permissions will need verification during authorized activation.

## Optional command-line use and maintenance

```powershell
py -3 market_report.py --as-of 2026-10-01
py -3 -m unittest discover -s tests -v
```

The program also accepts `--input`, `--output`, `--archive`, and `--settings` paths. By default Archive sits beside Output. Input, Output and Archive must be separate folders, with none inside another. Default paths are relative to the program's own folder, so launching it from another directory still works.

`settings.json` stores the file names, confirmed column mappings, status groups and community aliases. You do not need to edit it for routine monthly reports. If MRED changes a field or you add communities, have the mapping reviewed and update that file. Keep a copy of this whole project folder when transferring the system to another computer.

## Monthly history for the website

Each public JSON includes `monthly_history` as its 21st field. It contains exactly 12 complete calendar months, oldest first, including months without sales. Each object contains only `month`, `closed_sales`, `median_sold_price`, `sale_to_original_list_percentage`, and `median_closed_dom`. For October 1, 2026, the months are October 2025 through September 2026.

Monthly percentages use total Sold Pr divided by total Orig List Pr, multiplied by 100. Every qualifying original price must be positive and usable; regular List Price is never substituted. JSON preserves the existing four-decimal percentage precision; displayed percentages should use two decimal places. Zero-sales months contain a count of zero and three null statistics. Median sold price and DOM use that month's sales. Annual medians still use all qualifying annual sales directly.

Validation checks exact monthly fields, numeric values, consecutive chronological months, period boundaries, monthly sales totals against the annual count, and agreement with the latest-complete-month fields. Invalid history blocks report generation and GitHub publication. The monthly workflow requires a report date on the first day of a month so its annual period contains exactly 12 complete months.

The approved October enhancement reads the archived October source copies without changing them, because the newer Input exports differ from the approved snapshot. It only adds monthly_history to the seven current public JSON files. Existing 20 values, archived reports, audit records, and market_history.xlsx remain unchanged. Future monthly runs calculate history from their own validated input exports. GitHub publishing remains disabled and its monthly completion protections are unchanged.

## Inventory and closed-sales dates

Public JSON has exactly 23 fields. `inventory_updated` records the inventory refresh date; `closed_sales_through` records the inclusive closed-sales coverage cutoff. Both use YYYY-MM-DD. The cutoff equals `closed_period_end`, even when there were no sales on that day. It is not the last individual sale date.

For the current October baseline these dates are 2026-10-06 and 2026-09-30. Adding them does not recalculate the current reports or change any existing values.

For a daily refresh, replace only the Active/Pending export and double-click **Run Inventory.cmd**. Enter the inventory update date or press Enter for today. The command validates all seven JSON files and updates only their four inventory metrics and `inventory_updated`. It never reads the closed-sales export or changes the closed-sales dates, statistics, monthly history, report date, archives, or audit files. It does not publish to GitHub. Ask for a separate publication after reviewing the results. An update date older than the current inventory refresh is rejected.

For a monthly run, **Run Reports.cmd** also asks for the inventory source refresh date. Supply it when processing an export collected earlier. If left blank, the system preserves the existing date when the inventory export hash matches the previous monthly source; otherwise it uses today's processing date. It never derives the inventory date from the closed-sales cutoff. `closed_sales_through` is calculated automatically from the monthly report's final included day. Existing archive/history and GitHub publication protections remain in place.

Command-line equivalents:

```powershell
py -3 inventory_update.py --inventory-date 2026-10-06
py -3 market_report.py --as-of 2026-11-01 --inventory-date 2026-11-01
```

Neither example is an instruction to rerun the protected October archive. All current audit files, archived records, and historical workbook rows remain untouched by the date-field migration. Future newly generated Definitions worksheets document the dates.
