# Data equivalence analyzer

Checks that every row of an **old table** (usually an Excel general ledger, *hovedbok*) is also present in a
**new table** (usually BigQuery). It handles column names that don't match and new tables that have extra columns.
When rows are missing, it explains why: formatting problems (rounding, time zones, leading zeros, æ/ø/å encoding),
scope problems (a date outside the period, a month not loaded), and real differences.

Nothing is written back to Excel or BigQuery. Every query is a read-only `SELECT`.

## Files

| File | What it is |
|---|---|
| `hovedbok_compare.py` | The comparison engine (command line). |
| `hovedbok_gui.py` | A tkinter window around the two scripts. Start here if you don't want the command line. |
| `bq_download.py` | Downloads a whole BigQuery table or query result to CSV, with no row limit. |

Keep all three in the same folder.

## Installation

Python 3.9 or newer.

```sh
pip install pandas openpyxl
# for BigQuery:
pip install google-cloud-bigquery db-dtypes pydata-google-auth
# optional, makes reading large Excel files much faster:
pip install python-calamine
```

**Signing in to BigQuery:** if you have a gcloud login, the scripts use it. If not, the first run opens a browser
so you can sign in with your Google account, and the login is saved for later runs. You don't need gcloud.

To check the install, run the self-test:

```sh
python hovedbok_compare.py selftest
```

## Using the GUI

```sh
python hovedbok_gui.py
```

1. **Old table:** pick the Excel file and sheet, the first data row, and optionally the columns (`1-8,10` or `A-H,J`)
   and their names. Press **Load**.
2. **New table:** enter a BigQuery table (`project.dataset.table`) with an optional `WHERE` filter, or pick a CSV/Excel file.
   You can also set column types (`STRING`, `NUMERIC`, `DATE`, `DATE_LOCAL`, ...).
3. Use the buttons from left to right:
   - **Check the data**: preview both tables.
   - **Match columns**: get a suggested mapping, edit it by hand, then check values per column.
   - **Compare**: optionally set column roles (date, voucher, account, amount), then run the comparison and open the Excel report.

Settings are remembered between sessions in `~/.hovedbok_compare_gui.json`. **Save settings for the command line**
writes a `settings.json` so you can repeat the same check with `run --config`.

## Using the command line

```sh
# 1. Suggest which new column matches each old column -> mapping.csv (check and edit it)
python hovedbok_compare.py suggest --old Hovedbok.xlsx --new-bq myproj.finance.hovedbok

# 2. Quick check: are each old column's values found in its matched new column?
python hovedbok_compare.py check-values --old Hovedbok.xlsx --new-bq myproj.finance.hovedbok --mapping mapping.csv

# 3. Compare whole rows
python hovedbok_compare.py compare --old Hovedbok.xlsx --new-bq myproj.finance.hovedbok --mapping mapping.csv

# Or all three in one go (suggest runs only if there is no mapping yet)
python hovedbok_compare.py run --config settings.json
```

### Subcommands

| Command | Purpose |
|---|---|
| `suggest` | Suggests a column mapping. The score is 60% value overlap, 30% name similarity and 10% similar number of distinct values. |
| `check-values` | For each mapped pair, shows how many old values exist in the new column. |
| `compare` | Row containment. Duplicates count: an old row must appear in the new table at least as many times as in the old one. |
| `run` | Runs suggest, check-values and compare, reading each table only once. |
| `preview` | Prints the columns and sample values of both tables. |
| `list-columns` | Prints the old file's columns. Add `--save x.pkl` to cache the loaded table. |
| `kinds` | Lists every type of difference the diagnostics can report. |
| `selftest` | Builds a small test table with known differences and checks that all of them are found. |

### Common options

- **Old table:** `--old`, `--old-sheet`, `--old-start-row`, `--old-columns`, `--old-names "Konto;Tekst;Beløp"`,
  `--old-header-row`, `--old-table`
- **New table:** `--new FILE`, or `--new-bq PROJECT.DATASET.TABLE` / `--new-sql query.sql`, plus `--new-where`,
  `--bq-project`, `--cache bq_cache.pkl`
- **Normalisation:** `--decimals 2`, `--casefold`, `--new-types "konto=STRING; belop=NUMERIC; dato=DATE_LOCAL"`
- **Column selection:** `--columns-file`, `--skip-columns-file`
- **Compare:** `--date-col`, `--voucher-col`, `--account-col`, `--amount-col`, `--out-dir`, `--no-excel`
- **Automation:** `--json` prints one JSON result on stdout and sends the log to stderr. `--config settings.json` reads
  options from a file, and command-line options override it. `--fail-on-missing` exits with code 2 if anything is missing.

Run `python hovedbok_compare.py <command> --help` for the full list.

## Output

`compare` writes to `--out-dir` (default `hovedbok_compare_results/`):

- `summary.txt`: the overall result
- `results.xlsx`: everything in one workbook
- `missing_rows.csv`: old rows not found, with the likely cause and the column that most often breaks the match
- `totals_by_month.csv`: old vs new totals per month and account
- `recodes.csv`: values that were consistently replaced by another value
- `extra_in_new.csv`: rows that exist only in the new table

`check-values` writes `value_check.csv` and `values_not_in_new.csv`.

## How values are compared

Before comparing, both sides are normalised so that the same meaning gives the same text:

- Norwegian and international number formats: `1 234,50`, `1.234,50` and `1234.5`
- Dates in several formats: `30.09.2025`, `2025-09-30 00:00:00+00:00`, and Excel day numbers
- Booleans: `ja`/`nei`, `true`/`false`, `1`/`0`
- Long IDs keep every digit, because numbers are handled as high-precision decimals

`DATE_LOCAL` converts a UTC timestamp to Europe/Oslo time before taking the date.

## Downloading a BigQuery table

The BigQuery console stops downloads at a few MB. `bq_download.py` fetches the data page by page instead:

```sh
python bq_download.py --table myproj.finance.hovedbok --where "bilagsdato >= '2025-01-01'" --out hovedbok_2025.csv
python bq_download.py --sql query.sql --out result.csv --sep ";"
```
