#!/usr/bin/env python3
"""
hovedbok_compare.py - check whether every row of an old (Excel) table also
exists in a new (BigQuery) table, when the column names differ and the new
table has extra columns.

Two steps:

  1. suggest  - proposes which new column matches each old column, using name
                similarity and how many values the columns share. Writes a
                mapping CSV for you to check and edit.

  2. compare  - uses the mapping to compare rows. Every old row must appear in
                the new table at least as many times as it appears in the old
                one (duplicates are counted). Writes a summary and the rows
                that are missing, plus which column most often breaks a match.

  check-values - quick check per matched column pair: are the old column's
                non-empty values found in the new column? Good for checking a
                mapping before comparing whole rows.

  run         - all of it in one go: suggest (only if there is no mapping yet),
                check-values and compare, reading each table once.

  selftest    - makes a small test table with known differences and checks
                that they are all found. Run it once after installing.

Sources (old and new alike):
  --old FILE / --new FILE          .xlsx, .xls, .csv or .pkl (cache file)
  --old-sheet / --new-sheet NAME   sheet name for Excel (default: first sheet)
  --old-header-row N               row with the column headers (default: uses the
                                   Excel table on the sheet, or guesses the row)
  --old-table NAME                 Excel table to read if the sheet has several
  --new-bq PROJECT.DATASET.TABLE   read the new table straight from BigQuery
  --new-sql FILE.sql               or run your own SELECT from a .sql file
  --bq-project PROJECT             billing project for the query (optional)
  --new-where "CONDITION"          SQL WHERE condition to fetch only the relevant
                                   BigQuery rows, e.g. "bilagsdato >= '2025-01-01'"
  --cache FILE.pkl                 save the BigQuery result here and reuse it
                                   next time instead of querying again

Choosing columns:
  --columns-file FILE.txt          compare only these old columns (one per line)
  --skip-columns-file FILE.txt     leave these old columns out (one per line)
  list-columns --old FILE          print the old file's column names
  --old-names "Konto;Tekst;Beløp"   name the selected old columns in order (blank keeps the number)

Examples:
  python hovedbok_compare.py suggest --old Hovedbok.xlsx --new-bq myproj.finance.Hovedbok
  python hovedbok_compare.py compare --old Hovedbok.xlsx --new-bq myproj.finance.Hovedbok --mapping mapping.csv

For other programs and AI (see CLI_README.md):
  --json                           print one JSON result on stdout (text goes to stderr)
  --config settings.json           read options from a JSON file (command line wins)
  --fail-on-missing                compare/check-values exit with code 2 if anything is missing
  kinds                            list every difference type the diagnostics can report

Needs: pandas, openpyxl; for BigQuery also google-cloud-bigquery, db-dtypes and
pydata-google-auth (first run opens the browser to sign in; no gcloud needed).
Read-only: nothing is written to Excel or BigQuery.
"""

import argparse
import csv
import datetime as dt
import decimal
import difflib
import functools
import json
import os
import math
import re
import sys
import time
from collections import Counter
from pathlib import Path

decimal.getcontext().prec = 400  # room for long IDs and huge numbers without rounding errors

try:
    import pandas as pd
except ImportError:
    sys.exit("pandas is missing. Install with:  pip install pandas openpyxl")


# ---------------------------------------------------------------- loading ---

def find_excel_header(path, sheet, table_name=None):
    """Return (header_row, first_col, last_col, last_row) 1-based, or None.

    Uses an Excel table (Insert > Table) on the sheet if there is one, otherwise
    the first row in the top 100 that looks like a header row.
    """
    import openpyxl
    from openpyxl.utils import range_boundaries
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    tables = dict(ws.tables.items())
    if table_name:
        if table_name not in tables:
            sys.exit(f"No Excel table named '{table_name}' on sheet '{ws.title}'. Tables there: {list(tables) or 'none'}")
        tables = {table_name: tables[table_name]}
    if len(tables) == 1:
        name, ref = next(iter(tables.items()))
        ref = ref if isinstance(ref, str) else ref.ref
        min_col, min_row, max_col, max_row = range_boundaries(ref)
        print(f"{path.name}: using Excel table '{name}' ({ref}) on sheet '{ws.title}'", file=sys.stderr)
        return min_row, min_col, max_col, max_row
    if len(tables) > 1:
        print(f"Sheet '{ws.title}' has several Excel tables {list(tables)}; pick one with --old-table. "
              f"Guessing the header row instead.", file=sys.stderr)
    # No table: header row = first row whose filled cells are all text and is
    # as wide as the widest of the top rows.
    top = [list(r) for r in ws.iter_rows(min_row=1, max_row=100, values_only=True)]
    filled = [sum(v is not None and str(v).strip() != "" for v in r) for r in top]
    widest = max(filled, default=0)
    for i, r in enumerate(top):
        vals = [v for v in r if v is not None and str(v).strip() != ""]
        if vals and len(vals) >= 0.8 * widest and all(isinstance(v, str) for v in vals):
            print(f"{path.name}: guessed header row {i + 1} on sheet '{ws.title}' "
                  f"(use --old-header-row to override)", file=sys.stderr)
            return i + 1, None, None, None
    return None


def parse_column_numbers(text):
    """'1-6, 8,10' -> [1, 2, 3, 4, 5, 6, 8, 10]. Letters work too: 'A-F,H'."""
    def num(tok):
        tok = tok.strip()
        if tok.isdigit():
            return int(tok)
        if tok.isalpha():
            n = 0
            for ch in tok.upper():
                n = n * 26 + ord(ch) - 64
            return n
        sys.exit(f"Can't read column '{tok}'. Use numbers like 1-6,8 (or letters like A-F,H).")
    out = []
    for part in str(text).split(","):
        if not part.strip():
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            out.extend(range(num(a), num(b) + 1))
        else:
            out.append(num(part))
    if not out or min(out) < 1:
        sys.exit("Column numbers start at 1.")
    return sorted(set(out))


def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def excel_error_cells(path, sheet=None):
    """{(row, column): '#N/A'} for the error cells of a sheet (1-based). The calamine reader and pandas
    turn error values into blanks, which hides broken formulas; this puts them back. Scans the sheet's
    XML for t="e" cells, so it is fast and needs no extra memory. Returns {} if the file is not .xlsx/.xlsm."""
    import posixpath
    import zipfile
    import xml.etree.ElementTree as ET
    try:
        z = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError):
        return {}
    with z:
        ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
              "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
        try:
            sheets = ET.fromstring(z.read("xl/workbook.xml")).findall("m:sheets/m:sheet", ns)
            rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
        except (KeyError, ET.ParseError):
            return {}
        el = next((x for x in sheets if x.get("name") == sheet), None) if sheet else (sheets[0] if sheets else None)
        if el is None:
            return {}
        target = rels.get(el.get(f"{{{ns['r']}}}id"), "")
        name = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
        if name not in z.namelist():
            return {}
        out, tail = {}, ""
        cell_re = re.compile(r'<c\b[^>]*\br="([A-Z]+)(\d+)"[^>]*>(?:(?!</c>).)*?<v>([^<]*)</v>', re.S)
        with z.open(name) as f:
            while True:
                chunk = f.read(1 << 24)
                text = tail + chunk.decode("utf-8", errors="ignore") if chunk else tail
                cut = text.rfind("</c>") + 4 if chunk else len(text)
                body, tail = text[:cut], text[cut:]
                pos = body.find('t="e"')
                while pos != -1:
                    start = body.rfind("<c", 0, pos)
                    m = cell_re.match(body, start)
                    if m:
                        col = 0
                        for ch in m.group(1):
                            col = col * 26 + ord(ch) - 64
                        out[(int(m.group(2)), col)] = m.group(3)
                    pos = body.find('t="e"', pos + 5)
                if not chunk:
                    break
        return out


def read_excel_numbered(path, sheet, start_row, columns):
    """Read an Excel sheet from start_row down, without headers. Columns are named
    by their number (1 = column A). Uses the fast 'calamine' reader when installed."""
    import time
    t0 = time.time()
    try:
        import python_calamine  # noqa: F401
        engine = "calamine"
    except ImportError:
        engine = None
    if engine:
        df = pd.read_excel(path, sheet_name=sheet if sheet else 0, header=None,
                           skiprows=start_row - 1, engine="calamine")
        df.columns = range(1, len(df.columns) + 1)
        df.index = range(start_row, start_row + len(df))
        for (r, c), v in excel_error_cells(path, sheet).items():
            if r in df.index and c in df.columns:
                if df[c].dtype != object:
                    df[c] = df[c].astype(object)
                df.at[r, c] = v
    else:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb[sheet] if sheet else wb.worksheets[0]
        want = columns
        rows = []
        for row in ws.iter_rows(min_row=start_row, values_only=True):
            if want:
                rows.append([row[c - 1] if c - 1 < len(row) else None for c in want])
            else:
                rows.append(list(row))
        wb.close()
        width = max((len(r) for r in rows), default=0)
        df = pd.DataFrame(rows, columns=want if want else range(1, width + 1))
    if columns:
        missing = [c for c in columns if c not in df.columns]
        if missing:
            sys.exit(f"The sheet has no columns {missing}; it has {len(df.columns)} columns.")
        df = df[columns]
    df.index = range(start_row, start_row + len(df))  # index = Excel row number, for reports
    df.attrs["excel_rows"] = True
    df = df.dropna(axis=0, how="all")
    if not columns:
        df = df.dropna(axis=1, how="all")
    df.columns = [str(c) for c in df.columns]
    print(f"{Path(path).name}: read {len(df)} rows from row {start_row}, columns "
          f"{', '.join(f'{c} ({col_letter(int(c))})' for c in df.columns)} "
          f"in {time.time() - t0:.0f}s{'' if engine else ' (install python-calamine for faster reading)'}",
          file=sys.stderr)
    return df


def parse_names(text):
    """'Konto; Tekst; ; Beløp' -> ['Konto', 'Tekst', '', 'Beløp']. Uses ; if present, else commas."""
    sep = ";" if ";" in text else ","
    return [n.strip() for n in text.split(sep)]


def apply_names(df, names):
    """Rename the columns in order; a blank name keeps the column number."""
    names = parse_names(names) if isinstance(names, str) else list(names)
    while names and not names[-1]:
        names.pop()
    if len(names) > len(df.columns):
        sys.exit(f"{len(names)} column names given, but only {len(df.columns)} columns were read "
                 f"({', '.join(map(str, df.columns))}). Give one name per selected column.")
    new = [n or str(c) for n, c in zip(names, df.columns)] + [str(c) for c in df.columns[len(names):]]
    dupes = sorted({n for n in new if new.count(n) > 1})
    if dupes:
        sys.exit(f"Column names must be unique; used more than once: {dupes}")
    if names:
        print("Column names: " + ", ".join(f"{c} -> {n}" for c, n in zip(df.columns, new) if str(c) != n),
              file=sys.stderr)
    df.columns = new
    return df


def read_file(path, sheet=None, header_row=None, table=None, start_row=None, columns=None, names=None):
    df = _read_file_cols(path, sheet, header_row, table, start_row, columns)
    return apply_names(df, names) if names else df


# Within one 'run', each file is read only once (reading a big Excel sheet is the slow part).
_MEMO = None


def _memo(key, load):
    if _MEMO is None:
        return load()
    if key not in _MEMO:
        _MEMO[key] = load()
    else:
        print("(reusing the table already read)", file=sys.stderr)
    return _MEMO[key].copy(deep=False)


def _read_file_cols(path, sheet=None, header_row=None, table=None, start_row=None, columns=None):
    p = Path(path)
    stamp = p.stat().st_mtime if p.exists() else None
    return _memo(("file", str(p.resolve()), stamp, sheet, header_row, table, start_row, columns),
                 lambda: _read_file_cols_now(path, sheet, header_row, table, start_row, columns))


def _read_file_cols_now(path, sheet=None, header_row=None, table=None, start_row=None, columns=None):
    ext = Path(path).suffix.lower()
    if ext in (".xlsx", ".xlsm", ".xls") and header_row is None and table is None:
        # Numbered mode: no header detection, columns are 1, 2, 3 ...
        return read_excel_numbered(readable(path), sheet, start_row or 2,
                                   parse_column_numbers(columns) if columns else None)
    df = _read_file(path, sheet, header_row, table)
    df.columns = [str(c).strip() for c in df.columns]
    if columns:
        cols = list(df.columns)
        df = df[[cols[n - 1] for n in parse_column_numbers(columns) if n <= len(cols)]]
    looks_like_data = [c for c in df.columns
                       if c.startswith("Unnamed:") or re.fullmatch(r"[\d\s.,:/-]+", c) or re.search(r"\.\d+$", c)]
    numbered = ext == ".pkl" or all(c.isdigit() for c in df.columns)  # loaded earlier with numbered columns
    if len(df.columns) and not numbered and len(looks_like_data) >= max(2, len(df.columns) // 4):
        print(f"WARNING: {len(looks_like_data)} of {len(df.columns)} column names look like data, not headers "
              f"(e.g. {looks_like_data[:3]}). The header row is probably wrong. If the data is an Excel table, "
              f"leave the header row empty; otherwise give the row number where the column names are.",
              file=sys.stderr)
    return df


def read_column_list(path):
    """One old column name per line; blank lines and lines starting with # are ignored."""
    with open(path, encoding="utf-8-sig") as f:
        return [ln.strip() for ln in f if ln.strip() and not ln.lstrip().startswith("#")]


def select_old_columns(old, args):
    """Limit the old table to --columns-file / --skip-columns-file, if given."""
    cols = list(old.columns)
    if getattr(args, "columns_file", None):
        wanted = read_column_list(args.columns_file)
        unknown = [c for c in wanted if c not in cols]
        if unknown:
            print(f"warning: not in the old file, ignored: {unknown}", file=sys.stderr)
        cols = [c for c in cols if c in wanted]
    if getattr(args, "skip_columns_file", None):
        skip = set(read_column_list(args.skip_columns_file))
        cols = [c for c in cols if c not in skip]
    if not cols:
        sys.exit("No old columns left to compare after applying the column selection.")
    if len(cols) != len(old.columns):
        print(f"Using {len(cols)} of {len(old.columns)} old columns: {', '.join(cols)}", file=sys.stderr)
    return cols


def readable(path):
    """Return a path we can read. If the file is locked (open in Excel, OneDrive
    still syncing), read a temporary copy instead, or explain what to do."""
    path = Path(path)
    if not path.exists():
        sys.exit(f"File not found: {path}")
    try:
        with open(path, "rb"):
            return path
    except PermissionError:
        pass
    import shutil
    import tempfile
    copy = Path(tempfile.gettempdir()) / f"hovedbok_copy_{path.name}"
    try:
        shutil.copyfile(path, copy)
        print(f"{path.name} is locked (probably open in Excel), so a temporary copy is read instead.", file=sys.stderr)
        return copy
    except OSError:
        sys.exit(f"Windows won't let the script read {path.name} (permission denied).\n"
                 "Usually the file is open in Excel, or OneDrive is still syncing it. Close it in Excel and try again.\n"
                 "If it is a OneDrive file, right-click it and choose 'Always keep on this device', or save a copy "
                 "outside OneDrive and pick that copy.")


def _read_file(path, sheet=None, header_row=None, table=None):
    path = Path(path)
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xlsm", ".xls", ".csv"):
        path = readable(path)
    if ext in (".xlsx", ".xlsm", ".xls"):
        first_col = last_col = last_row = None
        if header_row is None and ext != ".xls":
            found = find_excel_header(path, sheet, table)
            if found:
                header_row, first_col, last_col, last_row = found
        header_row = header_row or 1
        df = pd.read_excel(
            path, sheet_name=sheet if sheet else 0, header=header_row - 1,
            usecols=list(range(first_col - 1, last_col)) if first_col else None,
            nrows=(last_row - header_row) if last_row else None,
        )
        errors = excel_error_cells(path, sheet) if ext != ".xls" else {}
        for (r, c), v in errors.items():
            i, j = r - header_row - 1, c - (first_col or 1)
            if 0 <= i < len(df) and 0 <= j < len(df.columns):
                if df.iloc[:, j].dtype != object:
                    df.isetitem(j, df.iloc[:, j].astype(object))
                df.iat[i, j] = v
        df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
        df = df.loc[:, [not str(c).startswith("Unnamed:") or df[c].notna().any() for c in df.columns]]
        print(f"{path.name}: header row {header_row}, columns: {', '.join(map(str, df.columns))}", file=sys.stderr)
        return df
    if ext == ".csv":
        # Sniff ; vs , (Norwegian Excel exports often use ;). Read as text.
        with open(path, encoding="utf-8-sig", errors="replace") as f:
            sep = ";" if f.readline().count(";") > 0 else ","
        return pd.read_csv(path, sep=sep, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    if ext == ".pkl":
        return pd.read_pickle(path)
    sys.exit(f"Don't know how to read {path} (use .xlsx, .xls, .csv or .pkl)")


def quote_bq(name):
    return "`" + name.replace("`", "") + "`"


def make_bq_client(bigquery, project):
    """Use a gcloud login if there is one, otherwise sign in through the browser."""
    import google.auth
    import google.auth.exceptions
    try:
        google.auth.default()
        return bigquery.Client(project=project)
    except google.auth.exceptions.DefaultCredentialsError:
        pass
    try:
        import pydata_google_auth
    except ImportError:
        sys.exit("No Google sign-in found. Install the browser sign-in helper with:\n"
                 "  python -m pip install pydata-google-auth\nand run again.")
    print("Opening your browser to sign in to Google (saved for next time) ...", file=sys.stderr)
    creds = pydata_google_auth.get_user_credentials(["https://www.googleapis.com/auth/bigquery"])
    if not project:
        sys.exit("Give the Google Cloud project with --project / --bq-project.")
    return bigquery.Client(project=project, credentials=creds)


def bq_sql(table=None, sql_file=None, columns=None, limit=None, where=None):
    if sql_file:
        return Path(sql_file).read_text(encoding="utf-8")
    cols = ", ".join(quote_bq(c) for c in columns) if columns else "*"
    sql = f"SELECT {cols} FROM {quote_bq(table)}"
    if where and where.strip():
        sql += f" WHERE {where.strip()}"
    if limit:
        sql += f" LIMIT {int(limit)}"
    return sql


def read_bigquery(table=None, sql_file=None, project=None, columns=None, limit=None, where=None):
    try:
        from google.cloud import bigquery
    except ImportError:
        sys.exit("BigQuery client is missing. Install with:\n"
                 "  python -m pip install google-cloud-bigquery db-dtypes pydata-google-auth")
    sql = bq_sql(table, sql_file, columns, limit, where)
    print(f"Querying BigQuery:\n  {sql.strip()[:300]}", file=sys.stderr)
    if not project and table and table.strip('`').count('.') == 2:
        project = table.strip('`').split('.')[0]
    client = make_bq_client(bigquery, project)
    return client.query(sql).to_dataframe()


def load_new(args, columns=None, limit=None):
    if args.new:
        return read_file(args.new, args.new_sheet, header_row=1)
    if not (args.new_bq or args.new_sql):
        sys.exit("Give the new table with --new FILE, --new-bq PROJECT.DATASET.TABLE or --new-sql FILE.sql")
    where = getattr(args, "new_where", None)
    # The cache remembers which table and filter it came from; a different one is fetched again.
    source_key = f"{args.new_bq or ''}\n{Path(args.new_sql).read_text(encoding='utf-8') if args.new_sql else ''}\n{where or ''}"
    key_file = Path(str(args.cache) + ".source.txt") if args.cache else None
    if args.cache and Path(args.cache).exists() and key_file.exists() \
            and key_file.read_text(encoding="utf-8") == source_key:
        df = pd.read_pickle(args.cache)
        missing = [c for c in (columns or []) if c not in df.columns]
        if not missing:
            print(f"Using cached BigQuery result {args.cache} ({len(df)} rows)", file=sys.stderr)
            return df
        print(f"Cache lacks columns {missing}; querying BigQuery again.", file=sys.stderr)
    elif args.cache and Path(args.cache).exists():
        print("The table or filter changed since the cache was saved; querying BigQuery again.", file=sys.stderr)
    df = _memo(("bq", source_key, tuple(columns or ()), limit),
               lambda: read_bigquery(args.new_bq, args.new_sql, args.bq_project, columns, limit, where))
    print(f"BigQuery returned {len(df)} rows.", file=sys.stderr)
    if args.cache and not limit:
        df.to_pickle(args.cache)
        key_file.write_text(source_key, encoding="utf-8")
        print(f"Saved BigQuery result to {args.cache}", file=sys.stderr)
    return df


# ---------------------------------------------------------- normalisation ---

NUM_RE = re.compile(r"^[+-]?(\d{1,3}([  .]\d{3})+|\d+)([.,]\d+)?$|^[+-]?[.,]\d+$")


SCI_RE = re.compile(r"^[+-]?\d+(?:[.,]\d+)?[eE][+-]?\d+$")


def to_decimal(s):
    """Parse '1 234,50', '1.234,50', '1234.5', '-12' etc. Return Decimal or None."""
    s = s.strip().replace(" ", " ")
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    if SCI_RE.match(s):  # 1.23457E+15, as Excel shows long numbers
        try:
            return decimal.Decimal(s.replace(",", "."))
        except decimal.InvalidOperation:
            return None
    if not NUM_RE.match(s):
        return None
    s = s.replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    elif s.count(".") > 1:
        s = s.replace(".", "")
    try:
        return decimal.Decimal(s)
    except decimal.InvalidOperation:
        return None


DATE_PATTERNS = [
    (re.compile(r"^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T]00:00(?::00(?:\.0+)?)?)?$"), (1, 2, 3)),
    (re.compile(r"^(\d{1,2})\.(\d{1,2})\.(\d{4})(?: 00:00(?::00)?)?$"), (3, 2, 1)),
    (re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4})(?: 00:00(?::00)?)?$"), (3, 2, 1)),  # dd/mm/yyyy
]

# Norwegian date with a time: 30.09.2026 13:00 or 30.09.2026 13:00:05
NO_DATETIME_RE = re.compile(r"^(\d{1,2})[./](\d{1,2})[./](\d{4})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?$")

# A date with a time and optional time zone: 2026-09-30 00:00:00+00:00, 2026-09-30T00:00:00Z, ... UTC
DATETIME_RE = re.compile(r"^(\d{4})-(\d{1,2})-(\d{1,2})[ T](\d{1,2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?"
                         r"\s*(?:Z|UTC|[+-]\d{2}(?::?\d{2})?)?$")


def norm_value(v, decimals, casefold):
    """Turn a cell into a comparable string. Same meaning -> same string."""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, (pd.Timestamp, dt.datetime)):
        if v.hour == v.minute == v.second == 0 and getattr(v, "microsecond", 0) == 0:
            return v.strftime("%Y-%m-%d")
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, (int, float, decimal.Decimal)) or hasattr(v, "dtype"):
        try:
            d = decimal.Decimal(str(v))
        except decimal.InvalidOperation:
            return str(v)
        if isinstance(v, float) and (math.isinf(v)):
            return str(v)
        return fmt_decimal(d, decimals)
    s = re.sub(r"\s+", " ", str(v)).strip()
    if s == "":
        return ""
    mt = DATETIME_RE.match(s)
    mt_no = None if mt else NO_DATETIME_RE.match(s)
    if mt or mt_no:  # e.g. '2026-09-30 00:00:00+00:00' or '30.09.2026 13:00': midnight -> just the date
        if mt:
            y, mo, d, h, mi, sec = (int(g or 0) for g in mt.groups()[:6])
        else:
            d, mo, y, h, mi, sec = (int(g or 0) for g in mt_no.groups()[:6])
        try:
            if h == mi == sec == 0:
                return dt.date(y, mo, d).isoformat()
            return dt.datetime(y, mo, d, h, mi, sec).strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            pass
    for rx, (y, m, d) in DATE_PATTERNS:
        mt = rx.match(s)
        if mt:
            try:
                return dt.date(int(mt.group(y)), int(mt.group(m)), int(mt.group(d))).isoformat()
            except ValueError:
                break
    num = to_decimal(s)
    if num is not None:
        return fmt_decimal(num, decimals)
    if s.lower() in ("true", "false"):
        return s.lower()
    return s.casefold() if casefold else s


def fmt_decimal(d, decimals):
    q = decimal.Decimal(1).scaleb(-decimals)
    if not d.is_finite():  # NaN / Infinity
        return str(d)
    try:  # high precision, so long IDs (e.g. 30 digits) keep every digit
        d = d.quantize(q, rounding=decimal.ROUND_HALF_UP, context=decimal.Context(prec=400))
    except decimal.InvalidOperation:
        d = decimal.Context(prec=400).normalize(d)
    if d == 0:
        d = abs(d)  # -0.00 -> 0.00
    s = format(d, "f")
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    return s


TYPE_ALIASES = {"STRING": "STRING", "TEXT": "STRING", "STR": "STRING",
                "INTEGER": "INTEGER", "INT": "INTEGER", "INT64": "INTEGER",
                "NUMERIC": "NUMERIC", "BIGNUMERIC": "NUMERIC", "DECIMAL": "NUMERIC", "FLOAT": "NUMERIC",
                "FLOAT64": "NUMERIC", "NUMBER": "NUMERIC",
                "DATE": "DATE", "DATETIME": "TIMESTAMP", "TIMESTAMP": "TIMESTAMP",
                "DATE_LOCAL": "DATE_LOCAL", "DATE_OSLO": "DATE_LOCAL", "LOCAL_DATE": "DATE_LOCAL",
                "BOOL": "BOOL", "BOOLEAN": "BOOL", "AUTO": None}


def parse_types(text):
    """'konto=STRING; belop=NUMERIC' -> {'konto': 'STRING', 'belop': 'NUMERIC'}"""
    out = {}
    for part in re.split(r"[;\n]", text or ""):
        if not part.strip():
            continue
        if "=" not in part:
            sys.exit(f"Column types: write column=TYPE, got '{part.strip()}'")
        col, typ = (x.strip() for x in part.rsplit("=", 1))
        if typ.upper() not in TYPE_ALIASES:
            sys.exit(f"Column types: unknown type '{typ}' for {col}. Use one of STRING, INTEGER, NUMERIC, DATE, "
                     f"DATE_LOCAL, TIMESTAMP, BOOL (or AUTO).")
        if TYPE_ALIASES[typ.upper()]:
            out[col] = TYPE_ALIASES[typ.upper()]
    return out


LOCAL_TZ = "Europe/Oslo"
TZ_TEXT_RE = re.compile(r"(Z|UTC|[+-]\d{2}:?\d{2})\s*$")


@functools.lru_cache(maxsize=200000)
def _local_date_from_text(s):
    try:
        ts = pd.Timestamp(s.replace(" UTC", "+00:00"))
    except (ValueError, TypeError):
        return None
    return ts.tz_convert(LOCAL_TZ).date().isoformat() if ts.tzinfo is not None else None


def _to_local_date(v):
    """A timestamp WITH a time zone (BigQuery TIMESTAMP, '...+00:00', '... UTC') -> its date in Norwegian
    time, e.g. 2026-09-29 22:00:00+00:00 -> 2026-09-30. Returns None when v has no time zone."""
    try:
        if isinstance(v, pd.Timestamp) or isinstance(v, dt.datetime):
            ts = pd.Timestamp(v)
        elif isinstance(v, str) and TZ_TEXT_RE.search(v.strip()):
            return _local_date_from_text(v.strip())
        else:
            return None
        if ts.tzinfo is None:
            return None
        return ts.tz_convert(LOCAL_TZ).date().isoformat()
    except (ValueError, TypeError):
        return None


def norm_typed(v, typ, decimals, casefold):
    """Normalise a value as a given BigQuery type. Values that do not fit the type are kept as text."""
    if typ is None:
        return norm_value(v, decimals, casefold)
    if typ == "STRING":
        if v is None:
            return ""
        try:
            if pd.isna(v):
                return ""
        except (TypeError, ValueError):
            pass
        if isinstance(v, float) and v.is_integer():
            v = int(v)  # 123.0 read from a number cell -> '123'
        s = re.sub(r"\s+", " ", str(v)).strip()
        return s.casefold() if casefold else s
    if typ == "DATE_LOCAL":
        local = _to_local_date(v)
        if local:
            return local
        typ = "DATE"  # no time zone on the value: it is already a local date/time
    n = norm_value(v, decimals, casefold)
    if not n:
        return n
    if typ in ("INTEGER", "NUMERIC"):
        num = _num(n)
        if num is None:
            return n
        return fmt_decimal(num, 0 if typ == "INTEGER" else decimals)
    if typ in ("DATE", "TIMESTAMP"):
        d, has_time = _date(n)
        if d is None:
            num = _num(n)
            if num is not None and _serial_to_date(num):  # Excel day number
                d, has_time = _serial_to_date(num), False
            else:
                return n
        return n if (typ == "TIMESTAMP" and has_time) else d.isoformat()
    if typ == "BOOL":
        low = n.casefold()
        if low in ("true", "1", "ja", "yes", "j", "y", "sann"):
            return "true"
        if low in ("false", "0", "nei", "no", "n", "usann"):
            return "false"
    return n


def norm_column(series, decimals, casefold, typ=None):
    if typ:
        def f(v):
            return norm_typed(v, typ, decimals, casefold)
    else:
        def f(v):
            return norm_value(v, decimals, casefold)
    vals = series.tolist()
    if len(vals) < 2000:
        return [f(v) for v in vals]
    # Ledger columns repeat a lot (accounts, dates, codes): normalise each distinct value once.
    # The key includes the type, so 1, 1.0 and True stay apart. Blanks are handled without the cache.
    empty, cache, out, na = f(None), {}, [], pd.NA
    room = 1_000_000
    for v in vals:
        if v is None or v is na or (v.__class__ is float and v != v):
            out.append(empty)
            continue
        try:
            k = (v.__class__, v)
            r = cache.get(k)
            if r is None:
                r = f(v)
                if room:
                    cache[k] = r
                    room -= 1
        except TypeError:  # unhashable value
            r = f(v)
        out.append(r)
    return out


# ---------------------------------------------------------------- suggest ---

def simplify_name(name):
    s = str(name).casefold()
    for a, b in (("æ", "ae"), ("ø", "o"), ("å", "a")):
        s = s.replace(a, b)
    return re.sub(r"[^a-z0-9]", "", s)


def score_pairs(old, new, sample, decimals, casefold, types=None):
    """Match score for every (old, new) column pair: 60% value overlap + 30% name similarity + 10% similar
    number of different values (so a column of 4 department codes doesn't match an amount column that
    happens to contain 100, 200, 300 and 400).
    Value overlap = share of the old column's different values (in `sample` random rows) found in the new
    column (in up to 10 x `sample` random rows). A column type given for a new column is used for both sides."""
    types = types or {}
    # Random rows, not the first ones: the tables are often sorted differently or cover different periods.
    if len(old) > sample:
        old = old.sample(sample, random_state=0)
    full_new = new
    if len(new) > 10 * sample:
        new = new.sample(10 * sample, random_state=0)

    def distinct(df, col, typ=None):
        vals = set(norm_column(df[col], decimals, casefold, typ))
        vals.discard("")
        return vals

    new_vals = {c: distinct(new, c, types.get(c)) for c in new.columns}
    old_vals_by_type = {(c, t): distinct(old, c, t) for c in old.columns for t in {None, *types.values()}}
    old_vals = {c: old_vals_by_type[(c, None)] for c in old.columns}
    def name_similarity(oc, nc):
        if str(oc).strip().isdigit():
            return 0.0  # an unnamed (numbered) column says nothing about the name
        return difflib.SequenceMatcher(None, simplify_name(oc), simplify_name(nc)).ratio()

    def card(ov, nv):
        return min(len(ov), len(nv)) / max(len(ov), len(nv)) if ov and nv else 0.0

    candidates = []
    for oc in old.columns:
        for nc in new.columns:
            name_sim = name_similarity(oc, nc)
            ov = old_vals_by_type[(oc, types.get(nc))]
            overlap = len(ov & new_vals[nc]) / len(ov) if ov else 0.0
            score = 0.6 * overlap + 0.3 * name_sim + 0.1 * card(ov, new_vals[nc])
            candidates.append((score, oc, nc, name_sim, overlap))
    if full_new is not new:
        # Columns with many different values (voucher numbers, amounts) look like poor matches in a sample of
        # the new rows. Recount the promising pairs against the whole new column.
        best = {}
        for c in candidates:
            if 0.05 <= c[4] < 0.9 and c[0] > best.get(c[1], (0,))[0]:
                best[c[1]] = c
        recount = {c[2] for c in best.values()}
        if recount:
            print(f"Recounting {len(recount)} column(s) against all {len(full_new)} new rows ...", file=sys.stderr)
        full_vals = {nc: distinct(full_new, nc, types.get(nc)) for nc in recount}
        for idx, (score, oc, nc, name_sim, overlap) in enumerate(candidates):
            ov = old_vals_by_type[(oc, types.get(nc))]
            if nc in full_vals and ov:
                overlap = len(ov & full_vals[nc]) / len(ov)
                candidates[idx] = (0.6 * overlap + 0.3 * name_sim + 0.1 * card(ov, new_vals[nc]),
                                   oc, nc, name_sim, overlap)
    return old_vals, candidates


def cmd_suggest(args):
    old = read_file(args.old, args.old_sheet, args.old_header_row, args.old_table, args.old_start_row, args.old_columns,
                    args.old_names)
    old = old[select_old_columns(old, args)]
    new = load_new(args, limit=args.sample if (args.new_bq and not args.new) else None)
    print(f"Old: {len(old)} rows, {len(old.columns)} columns. New (sample): {len(new)} rows, {len(new.columns)} columns.",
          file=sys.stderr)

    old_vals, candidates = score_pairs(old, new, args.sample, args.decimals, args.casefold,
                                       parse_types(args.new_types))

    # Greedy one-to-one assignment, best score first.
    candidates.sort(key=lambda c: -c[0])
    chosen, used_old, used_new = {}, set(), set()
    for score, oc, nc, name_sim, overlap in candidates:
        if oc in used_old or nc in used_new:
            continue
        chosen[oc] = (nc, name_sim, overlap, score)
        used_old.add(oc)
        used_new.add(nc)

    rows = []
    for oc in old.columns:
        nc, name_sim, overlap, score = chosen.get(oc, ("", 0, 0, 0))
        if not old_vals[oc]:
            conf = "empty column"
        elif overlap >= 0.9 and len(old_vals[oc]) >= 5:
            conf = "high"
        elif overlap >= 0.5 or name_sim >= 0.8:
            conf = "check"
        else:
            conf = "low - probably no match"
        # Only a likely match goes in new_column; a weak guess is kept aside so it isn't compared by accident.
        guess = ""
        if conf.startswith(("low", "empty")) and nc:
            nc, guess = "", nc
            conf += " (left unmapped; see best_guess)"
        rows.append([oc, nc, f"{name_sim:.2f}", f"{overlap:.2f}", len(old_vals[oc]), conf, guess])

    with open(args.out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["old_column", "new_column", "name_similarity", "value_overlap", "old_distinct_values", "confidence",
                    "best_guess"])
        w.writerows(rows)

    print(f"\nSuggested mapping (also written to {args.out}):\n")
    width = max(len(str(r[0])) for r in rows)
    for r in rows:
        print(f"  {str(r[0]):<{width}}  ->  {r[1] or '(none)':<30} names {r[2]}  values {r[3]}  [{r[5]}]")
    print("\nOpen the mapping file, fix any wrong new_column, and clear new_column for old columns that\n"
          "should not be compared (for example a column you add in Power Query). Then run 'compare'.")
    return {"command": "suggest", "mapping_file": str(Path(args.out).resolve()), "old_rows": len(old), "new_rows_sampled": len(new),
            "mapping": [{"old_column": str(r[0]), "new_column": r[1] or None, "name_similarity": float(r[2]),
                         "value_overlap": float(r[3]), "old_distinct_values": r[4], "confidence": r[5],
                         "best_guess": r[6] or None}
                        for r in rows]}


# ------------------------------------------------------------ diagnostics ---
# Ways two values can differ although they probably mean the same thing. Each kind
# says whether it looks like a formatting/export problem or a real difference.

KINDS = {
    "case": ("upper/lower case differs", "format", "tick 'Ignore upper/lower case'"),
    "spacing": ("spaces or punctuation differ", "format", "check for extra spaces, dashes or dots"),
    "zeros": ("leading zeros differ (e.g. 0123 vs 123)", "format", "one side stores the code as a number"),
    "rounding": ("numbers differ only by rounding", "format", "try fewer decimals"),
    "sign": ("number has the opposite sign", "format", "debit/credit sign convention"),
    "scale": ("number is 100 or 1000 times larger/smaller", "format", "øre vs kroner, or thousands"),
    "day_shift": ("date is one day off", "format", "time zone: a UTC timestamp shown in local time, or vice versa"),
    "day_month": ("day and month swapped", "format", "date read as mm.dd instead of dd.mm"),
    "time": ("same date, different time of day", "format", "one side has a time, the other only a date"),
    "cut": ("one text is cut off", "format", "text truncated in one of the systems"),
    "similar": ("text is almost the same (typo or rename)", "check", "edited or renamed in one system"),
    "blank_zero": ("0 in one table, blank in the other", "format", "zero stored as an empty cell"),
    "placeholder": ("blank written as a placeholder ('-', 'null', 'n/a')", "format", "different way of writing blank"),
    "encoding": ("æ/ø/å garbled (e.g. 'Ã¸' instead of 'ø')", "format", "file saved/read with the wrong character set"),
    "letters": ("æ/ø/å written as ae/o/a, or accents differ", "format", "letters replaced in one system"),
    "split": ("same total, but split into a different number of rows", "format",
              "one row in one table is several rows in the other (or the reverse); the amounts add up"),
    "two_cols": ("two columns differ at once", "check", "see likely_cause for the two columns"),
    "sparse": ("mostly empty row in Excel", "check", "probably a subtotal, heading or note line, not a posting"),
    "serial_date": ("date stored as an Excel day number (e.g. 45930 = 2025-09-30)", "format",
                    "the Excel cell is formatted as a number, not a date"),
    "bool_num": ("true/false vs 1/0", "format", "a yes/no field stored as a number in one table"),
    "precision": ("long number lost its last digits", "format",
                  "a long ID stored as a decimal number (Excel keeps only 15 digits), or shown as 1.23E+15"),
    "excel_error": ("Excel error value (#N/A, #REF!, #VALUE! ...)", "check", "a broken formula in the Excel sheet"),
    "out_of_period": ("date outside the new table's period", "scope",
                      "the new table (or your SQL filter) does not cover this date"),
    "no_month": ("the new table has no rows at all for this month", "scope",
                 "month not loaded into BigQuery, or filtered out"),
    "reversal": ("cancelled out by another missing row (+x and -x)", "check",
                 "a posting and its reversal; the new system may net these away"),
    "empty_old": ("empty in the old table", "check", "blank cell in Excel"),
    "empty_new": ("empty in the new table", "check", "blank in BigQuery"),
    "date_moved": ("date moved by more than a day", "explained",
                   "re-dated or moved to another period (correction, accrual)"),
    "recode": ("value consistently replaced by another value", "explained",
               "a code, account or text renamed in the new system (see recodes.csv)"),
    "text_edited": ("only the description/text differs", "explained", "text edited after the Excel copy was made"),
    "voucher_total": ("voucher has the same total per account in old and new", "explained",
                      "the voucher's lines were re-split or re-posted, but the voucher adds up to the same"),
    "zero_line": ("amount is 0", "explained", "zero-amount lines are often not transferred"),
    "cancelled_voucher": ("voucher cancels itself out on every account and is not in the new table", "explained",
                          "posted and reversed in the same voucher; often left out of a migration"),
    "opening_closing": ("opening/closing balance entry", "explained",
                        "opening balances and year-end closing entries are often created separately "
                        "in the new system"),
    "different": ("different value", "real", "probably a genuine difference"),
}


PLACEHOLDERS = {"-", "--", "null", "none", "nan", "n/a", "na", "#n/a", "(blank)", "(tom)", "tom", "."}


def _ascii(s):
    """Fold æ/ø/å and accents to plain letters: 'Varekjøp' -> 'varekjop'."""
    import unicodedata
    s = s.casefold()
    for a, b in (("æ", "ae"), ("ø", "o"), ("å", "a"), ("ß", "ss")):
        s = s.replace(a, b)
    s = unicodedata.normalize("NFKD", s)
    return re.sub(r"[\W_]+", "", "".join(c for c in s if not unicodedata.combining(c)))


def _unmangle(s):
    """Repair UTF-8 text that was read as Latin-1/cp1252: 'LÃ¸nn' -> 'Lønn'."""
    for enc in ("cp1252", "latin-1"):
        try:
            return s.encode(enc).decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    return s


def _serial_to_date(num):
    """Excel day number -> date (1 = 1900-01-01); only for plausible bookkeeping dates."""
    if 20000 <= num <= 80000:
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(num))
    return None


def _same_15_digits(a, b):
    """Long whole numbers that agree once the longer one is cut to the shorter one's digits:
    a float keeps 15-17 digits, Excel only 15 (12345678901234567 -> 12345678901234500) and may show 1.23457E+15."""
    if max(abs(a), abs(b)) < 10 ** 11 or a % 1 or b % 1:
        return False
    for x, y in ((a, b), (b, a)):
        k = len(x.normalize().as_tuple().digits)
        if 5 <= k <= 17 and len(y.normalize().as_tuple().digits) > k and any(decimal.Context(prec=k, rounding=r).plus(y) == x
                           for r in (decimal.ROUND_HALF_EVEN, decimal.ROUND_DOWN)):
            return True
    return False


def _alnum(s):
    return re.sub(r"[\W_]+", "", s.casefold())


def _num(s):
    try:
        d = decimal.Decimal(s)
    except (decimal.InvalidOperation, ValueError):
        return None
    return d if d.is_finite() else None


def _month_of(v):
    """(year, month) of a date, or of an accounting period written 202507 / 2025-07; else None."""
    d = _date(v)[0]
    if d:
        return d.year, d.month
    m = re.fullmatch(r"(19\d\d|20\d\d)-?(0[1-9]|1[0-3])", v)
    return (int(m[1]), int(m[2])) if m else None


@functools.lru_cache(maxsize=100000)  # a ledger has few distinct dates but millions of rows
def _date(s):
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})(?: (\d{2}):(\d{2}):(\d{2}))?$", s)
    if not m:
        return None, None
    try:
        return dt.date(int(m[1]), int(m[2]), int(m[3])), m[4] is not None
    except ValueError:
        return None, None


def classify_difference(o, n):
    """Why does normalised old value o differ from normalised new value n?"""
    if o == n:
        return None
    if {o, n} & {""} and ({o, n} - {""}) <= {"0"}:
        return "blank_zero"
    if not o or not n:
        if (o or n).strip().casefold() in PLACEHOLDERS:
            return "placeholder"
        return "empty_old" if not o else "empty_new"
    if {o, n} in ({"true", "1"}, {"false", "0"}):
        return "bool_num"
    no, nn = _num(o), _num(n)
    for num, other in ((no, n), (nn, o)):
        if num is not None and _date(other)[0] and _serial_to_date(num) == _date(other)[0]:
            return "serial_date"
    if no is not None and nn is not None:
        if no == nn:
            return "zeros"
        if _same_15_digits(no, nn):
            return "precision"
        if no == -nn:
            return "sign"
        # within 5 øre, or one side rounded to whole kroner
        if abs(no - nn) <= decimal.Decimal("0.05") or \
                ((no % 1 == 0 or nn % 1 == 0) and abs(no - nn) <= decimal.Decimal("0.5")):
            return "rounding"
        if nn and no and (no % 1 or nn % 1) and any(no == nn * f or no * f == nn for f in (100, 1000)):
            return "scale"
        return "different"
    do, to = _date(o)
    dn, tn = _date(n)
    if do and dn:
        if do == dn:
            return "time"
        if abs((do - dn).days) == 1:
            return "day_shift"
        if do.day <= 12 and (do.year, do.day, do.month) == (dn.year, dn.month, dn.day):
            return "day_month"
        if abs((do - dn).days) <= 62:
            return "date_moved"
        return "different"
    if o.casefold() == n.casefold():
        return "case"
    for a, b in ((o, n), (n, o)):
        try:
            if a.encode("latin-1").decode("utf-8") == b or a.encode("cp1252").decode("utf-8") == b:
                return "encoding"
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    ao, an = _alnum(o), _alnum(n)
    if ao and ao == an:
        return "spacing"
    if _ascii(o) == _ascii(n):
        return "letters"
    if ao and ao.lstrip("0") == an.lstrip("0") and ao.lstrip("0"):
        return "zeros"
    if len(o) >= 3 and len(n) >= 3 and (o.startswith(n) or n.startswith(o)):
        return "cut"
    if difflib.SequenceMatcher(None, o.casefold(), n.casefold()).ratio() >= 0.85:
        return "similar"
    return "different"


class FormatIndex:
    """Look up old values in a set of new values, allowing for format differences.

    Only the keys that the wanted old values could match are stored, so this stays fast
    and small on millions of new values.

    In a big column almost every date has a neighbour one day off, and most amounts have
    a twin with the opposite sign. Such matches are coincidence, so each kind is first
    tried on values that DO match exactly; kinds that 'find' something for many of them
    are switched off for this column."""

    def __init__(self, values, wanted, calibrate=500, scan=300000):
        self.values = values
        wanted = {o for o in wanted if o and o not in values}
        # Calibrate on new values shaped like the wanted ones (whole numbers, dates with day <= 12 ...):
        # a whole 300 has a 299.56 nearby far more often than a 299.56 has a 300.
        shapes = {self.shape(o) for o in wanted}
        per_shape = {}
        if wanted:
            for _, v in zip(range(scan), values):
                if v:
                    sh = self.shape(v)
                    if sh in shapes and len(per_shape.setdefault(sh, [])) < calibrate:
                        per_shape[sh].append(v)
        sample = [v for vs in per_shape.values() for v in vs]
        needed = set()
        for o in list(wanted) + sample:
            needed.update(self._candidates(o))
        kinds = {k[0] for k in needed}
        self.alt = {}
        if needed:
            for v in values:
                for key in self._keys(v, kinds):
                    if key in needed:
                        hits = self.alt.setdefault(key, [])
                        if len(hits) < 2 and v not in hits:
                            hits.append(v)
        self.disabled = set()  # (kind, shape) pairs that are coincidence in this column
        for sh, vs in per_shape.items():
            if len(vs) >= 30:
                chance = Counter(self._lookup(v)[0] for v in vs)
                self.disabled |= {(kind, sh) for kind, cnt in chance.items()
                                  if kind != "different" and cnt / len(vs) > 0.03}

    @staticmethod
    def shape(v):
        """A rough type of value, so coincidences are measured on values like the one looked up."""
        num = _num(v)
        if num is not None:
            return ("num", num == num.to_integral_value(), len(str(abs(int(num)))))
        d, _ = _date(v)
        if d:
            return ("date", d.day <= 12)
        return ("text", min(len(v), 30) // 6)

    @staticmethod
    def _keys(v, kinds):
        """Ways a new value v can be found (only the kinds some old value asks for)."""
        if "alnum" in kinds:
            a = _alnum(v)
            yield ("alnum", a)
            if a.lstrip("0") != a:
                yield ("alnum", a.lstrip("0"))
        if "bool" in kinds and v in ("true", "false", "1", "0"):
            yield ("bool", v in ("true", "1"))
        num = _num(v)
        if num is not None:
            if "date" in kinds and _serial_to_date(num):
                yield ("date", _serial_to_date(num))
            if "prec" in kinds and abs(num) >= 10 ** 14:
                yield ("prec", decimal.Context(prec=6).plus(num))
            if "abs" in kinds:
                yield ("abs", abs(num))
            if "round0" in kinds:
                yield ("round0", num.quantize(decimal.Decimal(1), rounding=decimal.ROUND_HALF_UP))
            if "scaled" in kinds and num:
                for f in (100, 1000):
                    yield ("scaled", num * f)
                    yield ("scaled", num / f)
            return
        d, _ = _date(v)
        if d:
            if "date" in kinds:
                yield ("date", d)
                for delta in (-1, 1):
                    yield ("date", d + dt.timedelta(days=delta))
                if d.day <= 12:
                    yield ("date", dt.date(d.year, d.day, d.month))
            if "serial" in kinds:
                yield ("serial", (d - dt.date(1899, 12, 30)).days)
            return
        if "ascii" in kinds:
            yield ("ascii", _ascii(_unmangle(v)))

    @staticmethod
    def _candidates(o):
        """Keys under which a format variant of old value o would be stored."""
        a = _alnum(o)
        out = [("alnum", a), ("alnum", a.lstrip("0"))]
        if o in ("true", "false", "1", "0"):
            out.append(("bool", o in ("true", "1")))
        num = _num(o)
        if num is not None:
            if _serial_to_date(num):
                out.append(("date", _serial_to_date(num)))
            if abs(num) >= 10 ** 14:
                out.append(("prec", decimal.Context(prec=6).plus(num)))
            out += [("abs", abs(num)), ("round0", num.quantize(decimal.Decimal(1), rounding=decimal.ROUND_HALF_UP)),
                    ("scaled", num)]
            return out
        d, _ = _date(o)
        if d:
            return out + [("date", d), ("serial", (d - dt.date(1899, 12, 30)).days)]
        return out + [("ascii", _ascii(_unmangle(o)))]

    def _lookup(self, o, disabled=()):
        sh = self.shape(o) if disabled else None
        for key in self._candidates(o):
            for n in self.alt.get(key, ()):
                if n == o:
                    continue
                kind = classify_difference(o, n)
                if kind and kind != "different" and (kind, sh) not in disabled:
                    return kind, n
        return "different", ""

    def explain(self, o):
        """None if o is in the set, else (kind, example new value) or ('different', '')."""
        if not o or o in self.values:
            return None
        return self._lookup(o, self.disabled)


class Nearest:
    """Closest value in a set: nearest number or date, or the most similar text."""

    def __init__(self, values):
        import bisect
        self.bisect = bisect
        self.values = values
        self.nums = sorted(n for n in (_num(v) for v in values) if n is not None)
        self.dates = sorted(d for d in (_date(v)[0] for v in values) if d)
        self.texts = None

    def _closest(self, seq, x):
        i = self.bisect.bisect_left(seq, x)
        near = [seq[j] for j in (i - 1, i) if 0 <= j < len(seq)]
        return min(near, key=lambda y: abs(y - x)) if near else None

    def find(self, v):
        num = _num(v)
        if num is not None:
            c = self._closest(self.nums, num)
            return "" if c is None else fmt_decimal(c, 10)
        d, _ = _date(v)
        if d:
            c = self._closest(self.dates, d)
            return c.isoformat() if c else ""
        if self.texts is None:
            self.texts = [t for t in self.values if _num(t) is None] if len(self.values) <= 20000 else []
        m = difflib.get_close_matches(v, self.texts, n=1, cutoff=0.6)
        return m[0] if m else ""


def cell_text(v):
    if v is None:
        return ""
    if isinstance(v, (pd.Timestamp, dt.datetime)):
        return str(v)
    return str(v)[:80]


def describe_kind(kind):
    text, nature, hint = KINDS[kind]
    tag = {"format": "FORMAT", "check": "CHECK", "real": "REAL", "scope": "SCOPE", "explained": "EXPLAINED"}[nature]
    return f"[{tag}] {text}" + (f" - {hint}" if nature != "real" else "")


# ---------------------------------------------------------------- compare ---

def read_mapping(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        sample = f.read(2048)
        f.seek(0)
        delim = ";" if sample.count(";") >= sample.count(",") else ","
        mapping = [(r["old_column"], r["new_column"].strip()) for r in csv.DictReader(f, delimiter=delim)
                   if r.get("new_column", "").strip()]
    if not mapping:
        sys.exit(f"No mapped columns in {path}")
    return mapping


RESULTS_GUIDE = """\
HOW TO READ AND USE THESE RESULTS
=================================

What was checked
----------------
Every row of the old table (the Excel sheet) was looked up in the new table
(BigQuery), using only the columns listed in mapping.csv. A row counts as found when
a new row has the same values in all those columns. Before comparing, values are put
in one standard form: numbers are rounded to the chosen decimals ("1 000,50", 1000.5
and "1000.50" are the same), dates are compared as dates (30.09.2026 = 2026-09-30),
and extra spaces are ignored. The new table may have extra rows and columns; only
old rows that are missing count as a problem.

The files
---------
results.xlsx             Everything below in one Excel workbook: the summary first,
                         then one sheet per file, with filters on the header row.
summary.txt              The overview: how many old rows were found, how many were
                         not, and the likely causes. Start here.
missing_rows.csv         One line per old row that was not found (open in Excel; it
                         uses ; as separator). Columns:
                           excel_rows      where the row is in the Excel sheet (first 5)
                           times_in_old / times_in_new  how often the row occurs
                           category        FORMAT, CHECK, REAL or SCOPE (see below)
                           difference_type short code for the cause (see below)
                           likely_cause    the cause in words, with the values
                           voucher_status  whole voucher missing, or only some lines
                           closest_match_pct  share of the compared columns that agree
                                           with the most similar new row (e.g. 95 =
                                           19 of 20 columns); 0 = no similar row
                           closest_match_columns_differing
                                           how many columns differ, e.g. "1 of 19"
                           closest_match_differs_in / closest_new_values
                                           which columns differ from that row, and
                                           its values ('old' -> 'new')
                           nearest_match_differs_in / old_value / new_value
                                           the closest new row and where it differs
totals_by_month.csv      Sum of the amount per month (and account) in old and new,
                         and the amount on the old rows that were not found.
extra_in_new.csv         The other direction: new rows that are not in the old
                         table, counting only the old table's period (and
                         accounts). "more copies than in old" can mean rows loaded
                         twice into BigQuery; "not in old" is often just postings
                         made after the Excel copy was taken.
value_check.csv          (from Check values) per column pair: share of old values
                         that exist in the new column, with likely causes.
values_not_in_new.csv    (from Check values) every old value that the new column
                         lacks, where it is in Excel, the likely cause and the
                         closest new value.

The four categories
-------------------
FORMAT  The data is there, but written differently (case, rounding, sign, date one
        day off, Excel day number, cut-off text, split rows ...). Not missing data.
        Fix the format or the mapping and run again.
CHECK   Needs a person to look: blanks, near-identical text, duplicates, reversals
        (+x/-x), Excel errors (#N/A), subtotal lines, or rows where two or more
        columns differ.
REAL    A value differs or does not exist in the new table at all. Probably a
        genuine difference that must be explained.
SCOPE   Outside what the new table covers (date before/after its period, or a month
        with no rows in the new table). Not an error in the data, but check that
        the period is meant to be excluded.
EXPLAINED  A likely legitimate reason was found: a code or account consistently
        renamed (listed in recodes.csv), a date moved to another day or period,
        only the description edited, a voucher whose lines were re-split but whose
        total per account is the same, a zero-amount line, a voucher that was posted
        and reversed (0 on every account) and left out, or an opening/closing
        balance entry. Nothing to fix, but the
        business controller should confirm that the reason is acceptable.

FOR THE DATA MAINTAINER (who owns the data load and the comparison setup)
-------------------------------------------------------------------------
1. Read the "By column and type of difference" list in summary.txt. Work from the
   top: one column with many FORMAT rows is usually one fix.
2. Typical fixes:
   - case differs              -> tick "Ignore upper/lower case"
   - rounding                  -> lower "Decimals"
   - date one day off          -> time zone in the export; compare dates in the same
                                  zone (e.g. convert to local date in the SQL/view)
   - Excel day number          -> format the Excel column as a date
   - long IDs lost digits      -> store IDs as text in Excel, not as numbers
   - æ/ø/å garbled             -> export/read the CSV as UTF-8
   - sign / x100 (øre)         -> agree on one convention; adjust in the view
   - wrong column              -> fix mapping.csv (Edit by hand..., Check values)
   - a column that cannot match (e.g. added in Excel only) -> leave it out of the
     mapping
3. SCOPE rows: check the SQL filter and that all months are loaded into BigQuery.
4. Run the comparison again after each fix. What remains in REAL and CHECK goes to
   the business controller, together with missing_rows.csv (filter on category).
5. Keep summary.txt from each run; the numbers should go down run by run.

FOR THE BUSINESS CONTROLLER (who decides whether differences are acceptable)
----------------------------------------------------------------------------
1. Look at totals_by_month.csv first: column "difference" shows where old and new
   totals do not agree, and by how much. Small or zero differences with many
   FORMAT rows mean the money is the same, only written differently.
2. In missing_rows.csv, sort on closest_match_pct: rows at 90-95% are almost
   certainly the same posting with one changed value; rows at 0-50% have no
   similar row in the new table. Then filter category = REAL. For each, the likely_cause and
   voucher_status say what to look at:
   - "whole voucher not in new"  -> the posting was never loaded (or was deleted/
                                    moved). Look the voucher up in Unit4.
   - "x of y line(s) missing"    -> the voucher exists, but a line differs; check
                                    whether it was corrected after the Excel copy.
3. EXPLAINED rows: confirm the reasons group by group (filter on difference_type).
   For renamed codes, go through recodes.csv once: one decision covers all rows.
   CHECK rows: reversals (+x/-x) and duplicates are often legitimate (a posting and
   its correction, or a repeated posting). Excel errors and subtotal lines are
   problems in the old sheet, not in BigQuery.
4. Decide per group whether the difference is (a) explained and acceptable,
   (b) a correction needed in BigQuery/Unit4, or (c) an error in the old Excel
   sheet. Note the decision next to the rows in missing_rows.csv.
5. The comparison is done when every remaining row has an explanation.

Limits
------
- The program only compares the mapped columns. Differences in other columns are
  not seen.
- "Likely cause" is a best guess from patterns in the data, not a proof.
- Old rows are matched one-to-one: if a row is in Excel twice, it must be in the new
  table twice.
- Nothing is changed in Excel or BigQuery; the program only reads.
"""


def next_steps(diag, col_kinds, voucher_status, totals_written, mapping_warnings=()):
    """Short, result-specific advice for the end of summary.txt."""
    cats = Counter(c for c, _, _ in diag.values() if c)
    out = ["", "NEXT STEPS (see HOW_TO_READ_THE_RESULTS.txt for details)"]
    for w in mapping_warnings:
        out.append(f"Data maintainer: check the mapping first: {w}")
    undiagnosed = sum(1 for c, kind, _ in diag.values() if not c and not kind)
    if undiagnosed:
        out.append(f"Data maintainer: {undiagnosed} missing rows were not diagnosed. Fix the mapping/format issues "
                   f"above first; if many remain, raise --diagnose-limit (costs time).")
    fmt = [(oc, kind, n) for (oc, kind), n in col_kinds.most_common() if KINDS[kind][1] == "format"][:3]
    if fmt:
        out.append("Data maintainer: fix these format differences first, then run again:")
        for oc, kind, n in fmt:
            out.append(f"  - {oc}: {KINDS[kind][0]} ({n} rows) -> {KINDS[kind][2]}")
    if cats.get("SCOPE"):
        out.append(f"Data maintainer: {cats['SCOPE']} rows are outside the new table's period; check the SQL filter "
                   f"and that all months are loaded.")
    if mapping_warnings:
        out.append("Business controller: wait with the review until the mapping is fixed and the comparison is "
                   "run again; most of the missing rows are probably caused by the wrong column.")
    elif cats.get("REAL") or cats.get("CHECK"):
        whole = sum(1 for k, s in voucher_status.items()
                    if s.startswith("whole") and diag.get(k, ("",))[0] in ("REAL", "CHECK"))
        out.append(f"Business controller: review {cats.get('REAL', 0)} REAL and {cats.get('CHECK', 0)} CHECK rows "
                   f"in missing_rows.csv (filter on 'category')"
                   + (f"; {whole} of them belong to vouchers missing entirely from the new table" if whole else "")
                   + ".")
    if cats.get("EXPLAINED") and not mapping_warnings:
        out.append(f"Business controller: confirm the {cats['EXPLAINED']} EXPLAINED rows (legitimate reasons such as "
                   f"renamed codes or moved dates; recodes.csv lists the renamings).")
    if totals_written:
        out.append("Business controller: totals_by_month.csv shows where the amounts differ, month by month.")
    if len(out) == 2:
        out.append("Nothing to do: every old row was found in the new table.")
    return out


def write_guide(out):
    (Path(out) / "HOW_TO_READ_THE_RESULTS.txt").write_text(RESULTS_GUIDE, encoding="utf-8-sig")


EXCEL_ERRORS = {"#n/a", "#ref!", "#value!", "#div/0!", "#name?", "#num!", "#null!", "#getting_data", "#spill!",
                "#calc!", "#i/t", "#ref", "#verdi!", "#navn?", "#tall!", "#null"}

ROLE_WORDS = {
    "date": ("bilagsdato", "dato", "date", "posteringsdato", "transdate", "voucherdate", "docdate"),
    "voucher": ("bilagsnr", "bilag", "voucherno", "voucher", "transno", "dokument", "docno"),
    "account": ("hovedbokskonto", "konto", "account", "acc", "gl"),
    "amount": ("belop", "beloep", "amount", "sum", "verdi", "value", "amt"),
}


def detect_roles(mapping, sample_keys, given):
    """Which mapped column is the date, voucher, account and amount? Uses names first, then values.
    given: {role: old column name or None}. Returns {role: column index or None}."""
    names = [(simplify_name(o), simplify_name(n)) for o, n in mapping]
    cols = list(zip(*sample_keys)) if sample_keys else [[] for _ in mapping]

    def share(i, test):
        vals = [v for v in cols[i] if v]
        return sum(1 for v in vals if test(v)) / len(vals) if vals else 0

    is_date = [share(i, lambda v: _date(v)[0] is not None) >= 0.9 for i in range(len(mapping))]
    is_num = [share(i, lambda v: _num(v) is not None) >= 0.95 for i in range(len(mapping))]
    has_dec = [share(i, lambda v: _num(v) is not None and _num(v) % 1 != 0) >= 0.05 for i in range(len(mapping))]
    roles, used = {}, set()
    for role in ("date", "amount", "voucher", "account"):
        if given.get(role):
            idx = [i for i, (o, _) in enumerate(mapping) if o == given[role]]
            if not idx:
                print(f"warning: --{role}-col '{given[role]}' is not a mapped old column; guessing instead",
                      file=sys.stderr)
            else:
                roles[role] = idx[0]
                used.add(idx[0])
                continue
        fits = {"date": lambda i: is_date[i], "amount": lambda i: is_num[i] and has_dec[i],
                "voucher": lambda i: not is_date[i] and not has_dec[i],
                "account": lambda i: not is_date[i] and not has_dec[i]}[role]
        best = None
        for w in ROLE_WORDS[role]:
            for i, (on, nn) in enumerate(names):
                if i not in used and fits(i) and (w in on or w in nn):
                    best = i
                    break
            if best is not None:
                break
        if best is None and role in ("date", "amount"):
            cands = [i for i in range(len(mapping)) if i not in used and fits(i)]
            if role == "amount":
                cands.sort(key=lambda i: -len(set(cols[i])))
            best = cands[0] if cands else None
        roles[role] = best
        if best is not None:
            used.add(best)
    return roles


def closest_rows(missing, new_keys, cap=300, top=25):
    """For each missing old row, the new row that agrees in the most columns.

    Candidates come from the old row's rarer values (a value shared by more than `cap`
    new rows, like a common date, is too unspecific to point at a row); the best
    candidates are then compared column by column. Returns {old key: (equal columns, new key)}."""
    if not missing:
        return {}
    ncols = len(missing[0])
    needed = [{k[i] for k in missing} for i in range(ncols)]
    postings = [dict() for _ in range(ncols)]  # value -> list of new row numbers (None = too common)
    for r, k in enumerate(new_keys):
        for i in range(ncols):
            v = k[i]
            if v in needed[i]:
                lst = postings[i].get(v, [])
                if lst is None:
                    continue
                if len(lst) >= cap:
                    postings[i][v] = None
                else:
                    lst.append(r)
                    postings[i][v] = lst
    result = {}
    for k in missing:
        votes = Counter()
        for i in range(ncols):
            lst = postings[i].get(k[i])
            if lst:
                votes.update(lst)
        best = None
        for r, _ in votes.most_common(top):
            nk = new_keys[r]
            eq = sum(1 for a, b in zip(k, nk) if a == b)
            if best is None or eq > best[0]:
                best = (eq, nk)
        if best:
            result[k] = best
    return result


def cmd_compare(args):
    mapping = read_mapping(args.mapping)
    old = read_file(args.old, args.old_sheet, args.old_header_row, args.old_table, args.old_start_row, args.old_columns,
                    args.old_names)
    if args.columns_file or args.skip_columns_file:
        keep = set(select_old_columns(old, args))
        dropped = [o for o, _ in mapping if o not in keep]
        mapping = [(o, n) for o, n in mapping if o in keep]
        if dropped:
            print(f"Left out by the column selection: {', '.join(dropped)}", file=sys.stderr)
        if not mapping:
            sys.exit("None of the selected columns has a new_column in the mapping.")
    old_cols = [o for o, _ in mapping]
    new_cols = [n for _, n in mapping]
    missing_old = [c for c in old_cols if c not in old.columns]
    if missing_old:
        sys.exit(f"Columns in mapping not found in old file: {missing_old}. If you renamed or changed "
                 f"the old columns, make the mapping again (Suggest automatically or Edit by hand...).")
    new = load_new(args, columns=new_cols)
    missing_new = [c for c in new_cols if c not in new.columns]
    if missing_new:
        sys.exit(f"Columns in mapping not found in new table: {missing_new}")

    print(f"Comparing {len(mapping)} column(s): old {len(old)} rows vs new {len(new)} rows ...", file=sys.stderr)

    types = parse_types(args.new_types)
    unknown = [c for c in types if c not in new_cols]
    if unknown:
        print(f"warning: column types given for columns that are not in the mapping: {unknown}", file=sys.stderr)
    if types:
        print("Column types: " + ", ".join(f"{c} = {t}" for c, t in types.items() if c in new_cols)
              + " (applied to the BigQuery column and the Excel column mapped to it)", file=sys.stderr)
    pair_types = [types.get(n) for n in new_cols]

    def keys(df, cols):
        normed = [norm_column(df[c], args.decimals, args.casefold, t) for c, t in zip(cols, pair_types)]
        return list(zip(*normed)) if normed else []

    old_keys = keys(old, old_cols)
    new_keys = keys(new, new_cols)
    old_count, new_count = Counter(old_keys), Counter(new_keys)
    roles = detect_roles(mapping, new_keys[:5000], {r: getattr(args, f"{r}_col", None)
                                                    for r in ("date", "voucher", "account", "amount")})
    D, V, A, M = roles["date"], roles["voucher"], roles["account"], roles["amount"]
    print("Column roles: " + ", ".join(f"{r} = {mapping[i][0] if i is not None else '(not found)'}"
                                       for r, i in roles.items())
          + "  (change with Column roles... or --date-col/--voucher-col/--account-col/--amount-col)",
          file=sys.stderr)

    fully_missing, short = [], []  # (key, old n, new n)
    for k, n in old_count.items():
        m = new_count.get(k, 0)
        if m == 0:
            fully_missing.append((k, n, m))
        elif m < n:
            short.append((k, n, m))
    matched_rows = sum(min(n, new_count.get(k, 0)) for k, n in old_count.items())
    missing_rows = len(old_keys) - matched_rows

    # A column whose old values are mostly absent from its new column is probably matched wrongly, and then
    # every row is "missing" for that reason alone. Say so before anything else.
    mapping_warnings = []
    if missing_rows > 0.2 * len(old_keys):
        for i, (oc, nc) in enumerate(mapping):
            new_vals_i = {k[i] for k in new_count}
            filled = [(k[i], n) for k, n in old_count.items() if k[i]]
            total = sum(n for _, n in filled)
            found = sum(n for v, n in filled if v in new_vals_i)
            if total >= 20 and found < 0.5 * total:
                mapping_warnings.append(f"{oc} -> {nc}: only {100 * found / total:.0f}% of the old rows have a value "
                                        f"that exists in {nc}. Probably the wrong column, or a format difference "
                                        f"(run Check values).")
        for w in mapping_warnings:
            print("WARNING: " + w, file=sys.stderr)

    # For missing rows: which single column, if ignored, would make them match?
    blame = Counter()
    blame_example = {}
    near, near_rank = {}, {}
    problem_keys = [k for k, _, _ in fully_missing]
    if problem_keys and len(mapping) > 1:
        sample = problem_keys[: args.diagnose_limit]
        # A near match agrees on every column but one, so it shares the value of the most varied column (or
        # of the second most varied one, when that is the column left out). Filtering the new rows on those
        # two columns first avoids building an index over millions of rows for every column.
        variety = sorted(range(len(mapping)), key=lambda j: -len({k[j] for k in sample}))
        filtered = {}
        for j in variety[:2]:
            wanted_j = {k[j] for k in sample}
            filtered[j] = [k for k in new_count if k[j] in wanted_j]
        for i, (oc, nc) in enumerate(mapping):
            j = variety[0] if variety[0] != i else variety[1]
            wanted = {k[:i] + k[i + 1:] for k in sample}
            index = {}
            for k in filtered[j]:
                r = k[:i] + k[i + 1:]
                if r in wanted and r not in index:
                    index[r] = k[i]
            for k in sample:
                hit = index.get(k[:i] + k[i + 1:])
                if hit is not None:
                    blame[oc] += 1
                    blame_example.setdefault(oc, (k[i], hit))
                    # Keep the most telling near match: a format difference beats a plain different value.
                    rank = {"format": 0, "explained": 1, "check": 1, "scope": 1, "real": 2}[KINDS[classify_difference(k[i], hit) or "different"][1]]
                    if k not in near or rank < near_rank[k]:
                        near[k], near_rank[k] = (oc, k[i], hit), rank

    to_check = set(problem_keys[: args.diagnose_limit])

    # Split/merged rows: with one amount column ignored, do the old and new rows add up to the same total?
    regrouped = {}            # key -> (old column, old rows, old total, new rows, new total)
    if to_check and len(mapping) > 1:
        for i, (oc, nc) in enumerate(mapping):
            sample_vals = [k[i] for k in list(new_count)[:2000] if k[i]]
            nums = [_num(v) for v in sample_vals]
            if not sample_vals or sum(x is not None for x in nums) < 0.95 * len(sample_vals) \
                    or sum(1 for x in nums if x is not None and x % 1) < 0.05 * len(sample_vals):
                continue  # not an amount column (needs numbers, some with decimals)
            needed = {k[:i] + k[i + 1:] for k in to_check}
            j = variety[0] if variety[0] != i else variety[1]
            wanted_j = {k[j] for k in to_check}
            sums = {}
            for side, keys_list in ((0, old_keys), (1, new_keys)):
                for k in keys_list:
                    if k[j] not in wanted_j:
                        continue
                    r = k[:i] + k[i + 1:]
                    if r in needed:
                        num = _num(k[i]) or decimal.Decimal(0)
                        acc = sums.setdefault(r, [0, decimal.Decimal(0), 0, decimal.Decimal(0)])
                        acc[2 * side] += 1
                        acc[2 * side + 1] += num
            for k in to_check:
                acc = sums.get(k[:i] + k[i + 1:])
                if acc and acc[2] and acc[0] != acc[2] and abs(acc[1] - acc[3]) < decimal.Decimal("0.01"):
                    regrouped.setdefault(k, (oc, acc[0], acc[1], acc[2], acc[3]))

    # Why is each missing row missing? Look at the near match, or at values the new column lacks.
    pending = []              # rows explained below by values the new columns lack

    diag = {}                 # key -> (category, difference type, cause text)
    key_kinds = {}            # key -> [(old column, kind, (old value, new value))]

    def note(k, col, kind, example):
        key_kinds.setdefault(k, []).append((col, kind, example))
    new_dates = sorted({d for d in (_date(k[D])[0] for k in new_count) if d}) if D is not None else []
    new_months = {(d.year, d.month) for d in new_dates}
    for k, n, m in fully_missing + short:
        errs = [i for i, v in enumerate(k) if v.casefold() in EXCEL_ERRORS]
        if errs:
            diag[k] = ("CHECK", "excel_error", "; ".join(f"{mapping[i][0]}: '{k[i]}'" for i in errs)
                       + f": {KINDS['excel_error'][2]}")
            for i in errs:
                note(k, mapping[i][0], "excel_error", (k[i], ""))
            continue
        d = _date(k[D])[0] if D is not None and new_dates else None
        if d and not (new_dates[0] <= d <= new_dates[-1]):
            diag[k] = ("SCOPE", "out_of_period", f"{mapping[D][0]}: {d} is outside the new table's dates "
                       f"{new_dates[0]} - {new_dates[-1]}")
            note(k, mapping[D][0], "out_of_period", (str(d), ""))
            continue
        if d and (d.year, d.month) not in new_months:
            diag[k] = ("SCOPE", "no_month", f"{mapping[D][0]}: the new table has no rows in {d:%Y-%m}")
            note(k, mapping[D][0], "no_month", (f"{d:%Y-%m}", ""))
            continue
        if m:
            diag[k] = ("CHECK", "duplicates", f"in old {n} times, in new only {m} times")
            continue
        if k not in to_check:
            diag[k] = ("", "", "not diagnosed (raise --diagnose-limit)")
            continue
        if sum(1 for v in k if not v) * 2 >= len(k) and len(k) > 2:
            diag[k] = ("CHECK", "sparse", f"{sum(1 for v in k if not v)} of {len(k)} compared columns are empty: "
                       f"{KINDS['sparse'][2]}")
            note(k, "(whole row)", "sparse", (" | ".join(v for v in k if v)[:60], ""))
            continue
        if k in regrouped:
            oc, n_old, t_old, n_new, t_new = regrouped[k]
            diag[k] = ("FORMAT", "split", f"{oc}: {n_old} old row(s) vs {n_new} new row(s) with the same other values, "
                       f"same total {fmt_decimal(t_old, 2)}")
            note(k, oc, "split", (f"{n_old} row(s)", f"{n_new} row(s), total {fmt_decimal(t_new, 2)}"))
            continue
        if k in near:
            oc, o, nv = near[k]
            kind = classify_difference(o, nv) or "different"
            note(k, oc, kind, (o, nv))
            diag[k] = (KINDS[kind][1].upper(), kind,
                       f"{oc}: old '{o}' vs new '{nv}': {KINDS[kind][0]}")
            continue
        pending.append(k)

    indexes = {}
    for i in range(len(mapping)):
        if pending:
            values = {k[i] for k in new_count}
            if any(k[i] and k[i] not in values for k in pending):
                indexes[i] = FormatIndex(values, {k[i] for k in pending})
    for k in pending:
        causes = []
        for i, (oc, nc) in enumerate(mapping):
            found = indexes[i].explain(k[i]) if i in indexes else None
            if found:
                kind, nv = found
                causes.append((oc, kind, k[i], nv))
                note(k, oc, kind, (k[i], nv))
        if causes:
            natures = {KINDS[kind][1] for _, kind, _, _ in causes}
            cat = "REAL" if "real" in natures else "CHECK" if "check" in natures else "FORMAT"
            text = "; ".join(f"{oc}: '{o}' not in new column" + (f" (new has '{nv}': {KINDS[kind][0]})" if nv else "")
                             for oc, kind, o, nv in causes)
            diag[k] = (cat, ",".join(sorted({kind for _, kind, _, _ in causes})), text)
        else:
            diag[k] = ("CHECK", "combination", "every value exists in the new table, but not together in one row "
                       "(three or more columns differ)")

    # Rows still unexplained: does a new row match when two columns are ignored?
    near2 = {}                # key -> [(column index, new value), (column index, new value)]
    combo = [k for k, d in diag.items() if d[1] == "combination"]
    if combo and len(mapping) > 2:
        print(f"Looking for near matches in two columns for {len(combo)} rows ...", file=sys.stderr)
        best = {}
        distinct_new = list(new_count)
        ncol = len(mapping)
        # A new row that differs from a row in only two columns agrees with it in at least one of any three
        # columns. So only new rows sharing a value in one of the three most selective columns are candidates
        # (one index per column pair over the whole new table took minutes for a handful of rows).
        wanted = [{k[c] for k in combo} for c in range(ncol)]
        cols3 = sorted(range(ncol), key=lambda c: sum(1 for k in distinct_new if k[c] in wanted[c]))[:3]
        by_val = {}
        for pos, k in enumerate(distinct_new):
            for c in cols3:
                if k[c] in wanted[c]:
                    by_val.setdefault((c, k[c]), []).append(pos)
        for k in combo:
            first = {}   # (i, j) -> values of the first new row (in table order) differing in exactly i and j
            for pos in sorted({p for c in cols3 for p in by_val.get((c, k[c]), ())}):
                nk = distinct_new[pos]
                d = []
                for c in range(ncol):
                    if k[c] != nk[c]:
                        d.append(c)
                        if len(d) > 2:
                            break
                if len(d) == 2 and (d[0], d[1]) not in first:
                    first[(d[0], d[1])] = (nk[d[0]], nk[d[1]])
            for (i, j), hit in sorted(first.items()):
                kinds = (classify_difference(k[i], hit[0]) or "different",
                         classify_difference(k[j], hit[1]) or "different")
                rank = max({"format": 0, "explained": 1, "check": 1, "scope": 1, "real": 2}[KINDS[x][1]] for x in kinds)
                if k not in best or rank < best[k][0]:
                    best[k] = (rank, (i, j), hit, kinds)
        for k, (rank, (i, j), hit, kinds) in best.items():
            parts = [f"{mapping[c][0]}: old '{k[c]}' vs new '{h}' ({KINDS[x][0]})"
                     for c, h, x in ((i, hit[0], kinds[0]), (j, hit[1], kinds[1]))]
            for c, h, x in ((i, hit[0], kinds[0]), (j, hit[1], kinds[1])):
                note(k, mapping[c][0], x, (k[c], h))
            diag[k] = ("FORMAT" if rank == 0 else "CHECK", "two_cols:" + ",".join(kinds), "; ".join(parts))
            near2[k] = [(i, hit[0]), (j, hit[1])]

    # Closest new row for each missing row: how many columns agree?
    closest = {}
    if to_check:
        print(f"Finding the closest new row for {len(to_check)} missing rows ...", file=sys.stderr)
        closest = closest_rows(sorted(to_check), list(new_count))

    # Postings cancelled by a reversal: missing rows on the same voucher/account (or with all other values
    # equal) whose amounts add up to zero.
    if M is not None:
        group_cols = [c for c in (V, A) if c is not None] or [i for i in range(len(mapping)) if i not in (M, D)]
        groups = {}
        for k, n, m in fully_missing:
            if diag.get(k, ("",))[0] in ("SCOPE", "FORMAT"):
                continue
            num = _num(k[M])
            if num is not None and num != 0:
                groups.setdefault(tuple(k[c] for c in group_cols), []).append((k, num * n))
        for members in groups.values():
            if len(members) >= 2 and sum(x for _, x in members) == 0:
                for k, _ in members:
                    key_kinds.pop(k, None)
                    diag[k] = ("CHECK", "reversal", f"{len(members)} missing rows on "
                               f"{' / '.join(mapping[c][0] + ' ' + k[c] for c in group_cols[:2])} add up to 0: "
                               f"{KINDS['reversal'][2]}")
        # A reversal is often posted on a voucher of its own: then the pair shares everything except the
        # voucher number and the sign of the amount.
        if V is not None:
            others = [i for i in range(len(mapping)) if i not in (V, M)]
            pairs_ = {}
            for k, n, m in fully_missing:
                if diag.get(k, ("",))[0] in ("SCOPE", "FORMAT") or diag.get(k, ("", ""))[1] == "reversal":
                    continue
                num = _num(k[M])
                if num:
                    pairs_.setdefault(tuple(k[c] for c in others), []).append((k, num * n))
            for members in pairs_.values():
                if len(members) >= 2 and sum(x for _, x in members) == 0 \
                        and len({k[V] for k, _ in members}) > 1:
                    for k, _ in members:
                        key_kinds.pop(k, None)
                        diag[k] = ("CHECK", "reversal", f"{len(members)} missing rows on different "
                                   f"{mapping[V][0]} ({', '.join(sorted({m[V] for m, _ in members})[:3])}) are equal "
                                   f"apart from opposite amounts that add up to 0: {KINDS['reversal'][2]}")
                    note(k, mapping[M][0], "reversal", (k[M], ""))

    # ---- Legitimate reasons: differences that are explained by how the data was moved, not errors.
    legit = Counter()
    open_rows = [k for k, n, m in fully_missing if diag.get(k, ("",))[0] in ("CHECK", "REAL")
                 and diag.get(k, ("", ""))[1] not in ("excel_error", "sparse")]  # broken/odd Excel rows stay CHECK

    def explain(k, kind, text, col=None, example=("", "")):
        key_kinds.pop(k, None)
        diag[k] = ("EXPLAINED", kind, text)
        note(k, col or "(whole row)", kind, example)
        legit[kind] += 1

    # 1. Zero-amount lines.
    if M is not None:
        for k in open_rows:
            if _num(k[M]) == 0:
                explain(k, "zero_line", f"{mapping[M][0]} is 0: {KINDS['zero_line'][2]}", mapping[M][0], (k[M], ""))

    # 2. Opening balance / year-end closing entries (recognised from the text in any column).
    ob_re = re.compile(r"\b(ib|ub|inng[aå]ende (?:balanse|saldo)|[aå]pningsbalanse|opening balance|utg[aå]ende balanse|"
                       r"[aå]rsavslutning|[aå]rsoppgj[oø]r|resultatdisponering|closing|year[- ]end)\b", re.I)
    for k in open_rows:
        if diag[k][0] in ("CHECK", "REAL") and any(ob_re.search(v) for v in k if v and _num(v) is None):
            explain(k, "opening_closing", KINDS["opening_closing"][2])

    # 3. Systematic recodes: an old value that, in the closest matching rows, is always the same new value.
    pairs, per_old = Counter(), Counter()
    for k in open_rows:
        diffs = []
        if k in near:
            oc, o, nv = near[k]
            diffs = [(old_cols.index(oc), nv)]
        elif k in near2:
            diffs = near2[k]
        for i, nv in diffs:
            if k[i] and nv:
                pairs[(i, k[i], nv)] += 1
                per_old[(i, k[i])] += 1
    recodes = {(i, o): (nv, cnt) for (i, o, nv), cnt in pairs.items()
               if cnt >= 3 and cnt >= 0.8 * per_old[(i, o)] and i not in (M, D)}
    # Several old values all "renamed" to one new value is not a renaming: usually the column is matched to
    # the wrong new column (e.g. five people -> one system name). Those stay unexplained.
    sources = Counter((i, nv) for (i, _), (nv, _) in recodes.items())
    recodes = {key: val for key, val in recodes.items() if sources[(key[0], val[0])] < 3}
    if recodes:
        Path(args.out_dir).mkdir(parents=True, exist_ok=True)
        with open(Path(args.out_dir) / "recodes.csv", "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["old_column", "old_value", "new_value", "rows"])
            for (i, o), (nv, cnt) in sorted(recodes.items(), key=lambda x: -x[1][1]):
                w.writerow([mapping[i][0], o, nv, cnt])
    for k in open_rows:
        if diag[k][0] not in ("CHECK", "REAL"):
            continue
        diffs = [(old_cols.index(near[k][0]), near[k][2])] if k in near else near2.get(k, [])
        if diffs and all((i, k[i]) in recodes and recodes[(i, k[i])][0] == nv for i, nv in diffs):
            i, nv = diffs[0]
            explain(k, "recode", "; ".join(f"{mapping[i][0]}: '{k[i]}' is consistently '{nv}' in the new table "
                                           f"({recodes[(i, k[i])][1]} rows)" for i, nv in diffs),
                    mapping[i][0], (k[i], nv))

    # 4. Only a description changed, or the date moved within two months.
    text_cols = set()
    for i in range(len(mapping)):
        vals = [k[i] for k in list(new_count)[:3000] if k[i]]
        if vals and sum(1 for v in vals if _num(v) is None and _date(v)[0] is None) >= 0.9 * len(vals) \
                and len(set(vals)) >= 20:
            text_cols.add(i)
    for k in open_rows:
        if diag[k][0] not in ("CHECK", "REAL"):
            continue
        diffs = [(old_cols.index(near[k][0]), near[k][2])] if k in near else near2.get(k, [])
        if not diffs:
            continue
        kinds_ = [classify_difference(k[i], nv) or "different" for i, nv in diffs]
        ok = [i in text_cols or kd == "date_moved" or KINDS[kd][1] == "format" for (i, _), kd in zip(diffs, kinds_)]
        if all(ok) and any(i in text_cols for i, _ in diffs) and "similar" not in kinds_:
            # (a one-letter change stays CHECK 'similar': it may be a typo, not an edit)
            i = next(i for i, _ in diffs if i in text_cols)
            explain(k, "text_edited", "; ".join(f"{mapping[i][0]}: '{k[i]}' -> '{nv}'" for i, nv in diffs)
                    + f": {KINDS['text_edited'][2]}", mapping[i][0], (k[i], dict(diffs)[i]))
        elif all(ok) and "date_moved" in kinds_:
            i, nv = diffs[kinds_.index("date_moved")]
            days = (_date(nv)[0] - _date(k[i])[0]).days
            explain(k, "date_moved", f"{mapping[i][0]}: '{k[i]}' -> '{nv}' ({days:+d} days): "
                    f"{KINDS['date_moved'][2]}", mapping[i][0], (k[i], nv))

    # 5. Vouchers: same total in old and new (lines re-split), or a voucher cancelled by its own reversal.
    # A complete voucher always adds up to 0 (debit = credit), so totals are taken per voucher AND account:
    # re-split lines keep the account's total, and a cancelled voucher nets to 0 on every account.
    if V is not None and M is not None:
        vt_old, vt_new, vn_old = {}, {}, Counter()
        wanted_v = {k[V] for k in open_rows if k[V]}

        def vkey(k):
            return (k[V], k[A]) if A is not None else (k[V],)
        for counter, totals_ in ((old_count, vt_old), (new_count, vt_new)):
            for k, n in counter.items():
                if k[V] in wanted_v:
                    totals_[vkey(k)] = totals_.get(vkey(k), decimal.Decimal(0)) + (_num(k[M]) or 0) * n
                    if counter is old_count:
                        vn_old[k[V]] += n
        new_vouchers = {vk[0] for vk in vt_new}
        old_by_voucher = {}
        for vk, t in vt_old.items():
            old_by_voucher.setdefault(vk[0], []).append(t)
        for k in open_rows:
            if diag[k][0] not in ("CHECK", "REAL") or not k[V]:
                continue
            v, vk = k[V], vkey(k)
            what = f"{mapping[V][0]} {v}" + (f", {mapping[A][0]} {k[A]}" if A is not None else "")
            diffs = [old_cols.index(near[k][0])] if k in near else [i for i, _ in near2.get(k, [])]
            if (vk in vt_new and abs(vt_old[vk] - vt_new[vk]) < decimal.Decimal("0.01")
                    and (A is not None or vt_old[vk] != 0) and (not diffs or M in diffs)):
                explain(k, "voucher_total", f"{what}: total {fmt_decimal(vt_old[vk], 2)} in both old and "
                        f"new: {KINDS['voucher_total'][2]}", mapping[V][0], (v, ""))
            elif (A is not None and v not in new_vouchers and vn_old[v] >= 2
                  and all(t == 0 for t in old_by_voucher[v])):
                explain(k, "cancelled_voucher", f"{mapping[V][0]} {v} ({vn_old[v]} lines) adds up to 0 on every "
                        f"{mapping[A][0]} and is not in the new table: {KINDS['cancelled_voucher'][2]}",
                        mapping[V][0], (v, ""))

    col_kinds, col_example = Counter(), {}   # (old column, kind) -> missing rows, example
    for notes in key_kinds.values():
        for col, kind, example in notes:
            col_kinds[(col, kind)] += 1
            col_example.setdefault((col, kind), example)

    # Is the whole voucher missing, or only some of its lines?
    voucher_status = {}
    if V is not None:
        new_vouchers = {k[V] for k in new_count}
        old_lines = Counter()
        for k, n in old_count.items():
            old_lines[k[V]] += n
        missing_lines = Counter()
        for k, n, m in fully_missing + short:
            missing_lines[k[V]] += n - m
        for k, n, m in fully_missing + short:
            v = k[V]
            if not v:
                continue
            if v not in new_vouchers:
                voucher_status[k] = f"whole voucher not in new ({old_lines[v]} line(s))"
            elif missing_lines[v] >= old_lines[v]:
                voucher_status[k] = f"voucher in new, but none of its {old_lines[v]} old line(s) match"
            else:
                voucher_status[k] = f"{missing_lines[v]} of {old_lines[v]} line(s) missing"

    # Where each missing row sits in the Excel sheet, so it can be looked up there.
    missing_set = {k for k, _, _ in fully_missing + short}
    where = {}
    row_nos = old.index if old.attrs.get("excel_rows") else range(2, len(old) + 2)
    for k, r in zip(old_keys, row_nos):
        if k in missing_set:
            where.setdefault(k, []).append(r)

    out = Path(args.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    with open(out / "missing_rows.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["excel_rows"] + old_cols + ["times_in_old", "times_in_new", "category", "difference_type", "likely_cause",
                               "voucher_status", "closest_match_pct", "closest_match_columns_differing", "closest_match_differs_in",
                               "closest_new_values", "nearest_match_differs_in", "old_value", "new_value"])
        for k, n, m in sorted(fully_missing + short):
            hint = near.get(k, ("", "", ""))
            if k in closest:
                eq, nk = closest[k]
                diff = [i for i in range(len(k)) if k[i] != nk[i]]
                best_cols = [f"{(100 * eq / len(k)):.0f}", f"{len(diff)} of {len(k)}",
                             ", ".join(mapping[i][0] for i in diff),
                             " | ".join(f"{mapping[i][0]}: '{k[i]}' -> '{nk[i]}'" for i in diff)]
            else:
                best_cols = ["0" if m == 0 and k in to_check else "", "", "", ""]
            rows_at = where.get(k, [])
            w.writerow([", ".join(str(r) for r in rows_at[:5]) + (" ..." if len(rows_at) > 5 else "")]
                       + list(k) + [n, m, *diag.get(k, ("", "", "")), voucher_status.get(k, ""), *best_cols, *hint])

    # Totals per month (and account): where do the missing amounts sit?
    totals = {}
    if D is not None and M is not None:
        def bucket(k, old_side=True):
            d = _date(k[D])[0]
            acc = k[A] if A is not None else None
            if old_side and A is not None and (A, acc) in recodes:
                acc = recodes[(A, acc)][0]  # renamed account: total it under its new number
            return (f"{d:%Y-%m}" if d else "(no date)",) + ((acc,) if A is not None else ())
        for side, counter in ((0, old_count), (1, new_count)):
            for k, n in counter.items():
                b = bucket(k, side == 0)
                if side == 1 and b not in totals:
                    continue  # only months/accounts that are in the old table
                t = totals.setdefault(b, [decimal.Decimal(0), decimal.Decimal(0), decimal.Decimal(0), 0, 0])
                t[side] += (_num(k[M]) or 0) * n
                t[3 + side] += n
        for k, n, m in fully_missing + short:
            totals[bucket(k)][2] += (_num(k[M]) or 0) * (n - m)
        with open(out / "totals_by_month.csv", "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["month"] + (["account"] if A is not None else []) +
                       ["old_total", "new_total", "difference", "amount_of_missing_old_rows", "old_rows", "new_rows"])
            for b, t in sorted(totals.items()):
                w.writerow(list(b) + [fmt_decimal(t[0], 2), fmt_decimal(t[1], 2), fmt_decimal(t[1] - t[0], 2),
                                      fmt_decimal(t[2], 2), t[3], t[4]])

    # The other direction: new rows inside the old table's period (and accounts) that the old table lacks.
    # Rows occurring more often in new than in old are the strongest sign of a double load.
    extra, more_copies = [], 0
    if D is not None:
        # The old period, ignoring the odd stray row (a 2021 correction in a 2024-2026 sheet): 0.5% at each end.
        old_dates = sorted(d for k, n in old_count.items() for d in [_date(k[D])[0]] * n if d)
        lo = hi = None
        if old_dates:
            cut = len(old_dates) // 200
            q_lo, q_hi = old_dates[cut], old_dates[-1 - cut]
            month = dt.timedelta(days=31)
            lo = next(d for d in old_dates if d >= q_lo - month)          # drop only dates far before the rest
            hi = next(d for d in reversed(old_dates) if d <= q_hi + month)
        old_accounts = {k[A] for k in old_count} if A is not None else None
    for k, m in new_count.items():
        n = old_count.get(k, 0)
        if m <= n:
            continue
        if n == 0:
            if D is None:
                continue  # without a date we can't tell which new rows the old table should have had
            d = _date(k[D])[0]
            if not d or lo is None or not (lo <= d <= hi) or (old_accounts is not None and k[A] not in old_accounts):
                continue
        else:
            more_copies += m - n
        extra.append((k, m, n))
    extra_rows = sum(m - n for _, m, n in extra)
    # A new row that is the closest match of a missing old row is probably the same posting, changed.
    counterpart = {nk for eq, nk in closest.values() if nk is not None and eq >= len(mapping) - 2}
    paired = sum(m - n for k, m, n in extra if n == 0 and k in counterpart)
    if extra or (out / "extra_in_new.csv").exists():
        extra.sort(key=lambda e: (e[2] == 0, e[0]))
        cap = 200000
        with open(out / "extra_in_new.csv", "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(new_cols + ["times_in_new", "times_in_old", "reason"])
            for k, m, n in extra[:cap]:
                w.writerow(list(k) + [m, n, f"{m - n} more copies than in old" if n else
                                      "changed version of a missing old row (closest match in missing_rows.csv)"
                                      if k in counterpart else
                                      "not in old (inside the old period" + (" and accounts)" if A is not None else ")")])
        if len(extra) > cap:
            print(f"extra_in_new.csv lists the first {cap} of {len(extra)} rows.", file=sys.stderr)

    lines = [f"WARNING - check the mapping: {w}" for w in mapping_warnings] + ([""] if mapping_warnings else []) + [
        f"Old rows:              {len(old_keys)}",
        f"New rows:              {len(new_keys)}",
        f"Columns compared:      {len(mapping)}",
        f"Old rows found in new: {matched_rows}",
        f"Old rows NOT in new:   {missing_rows}"
        + (f"  ({len(fully_missing)} distinct rows absent, {len(short)} rows with fewer duplicates in new)"
           if missing_rows else ""),
        "",
        "RESULT: the old table IS a subset of the new one on the compared columns."
        if missing_rows == 0 else
        "RESULT: the old table is NOT fully contained in the new one. See missing_rows.csv.",
    ]
    if diag:
        cats = Counter(c for c, _, _ in diag.values() if c)
        lines += ["", "Likely causes for the missing rows (distinct rows; 'category' column in missing_rows.csv):",
                  f"  FORMAT  {cats.get('FORMAT', 0):>7}  same data, written differently - fix the format, not the data",
                  f"  CHECK   {cats.get('CHECK', 0):>7}  needs a look (blanks, near-identical text, duplicates, "
                  f"or values that exist but not together)",
                  f"  REAL    {cats.get('REAL', 0):>7}  a genuine difference: a value differs, or is not in the new "
                  f"table at all",
                  f"  SCOPE   {cats.get('SCOPE', 0):>7}  outside what the new table covers (date range or missing "
                  f"month) - not an error",
                  f"  EXPLAINED {cats.get('EXPLAINED', 0):>5}  a likely legitimate reason was found (renamed code, "
                  f"moved date, edited text, re-split voucher ...) - confirm, don't fix"]
        if legit:
            lines += ["", "Legitimate reasons found (category EXPLAINED):"]
            for kind, cnt in legit.most_common():
                lines.append(f"  {cnt:>7}  {KINDS[kind][0]} - {KINDS[kind][2]}")
            if recodes:
                lines.append(f"  The renamed values are listed in recodes.csv ({len(recodes)} old value(s)); "
                             f"check that each renaming is intended.")
        undiagnosed = sum(1 for c, kd, _ in diag.values() if not c and not kd)
        if undiagnosed:
            lines.append(f"  ({undiagnosed} more missing rows were not diagnosed; "
                         f"raise --diagnose-limit to include them)")
        combos = sum(1 for c in diag.values() if c[1] == "combination")
        two = sum(1 for c in diag.values() if c[1].startswith("two_cols"))
        dupes = sum(1 for c in diag.values() if c[1] == "duplicates")
        if col_kinds:
            lines += ["", "By column and type of difference (missing rows; one row can count in several columns):"]
            width = max(len(str(oc)) for oc, _ in col_kinds)
            for (oc, kind), cnt in col_kinds.most_common():
                o, nv = col_example[(oc, kind)]
                eg = f"old '{o}' vs new '{nv}'" if nv else f"old '{o}'"
                lines.append(f"  {str(oc):<{width}}  {cnt:>7}  {describe_kind(kind)}   e.g. {eg}")
        if two:
            lines.append(f"\n{two} missing row(s) match a new row except in two columns (difference_type "
                         f"'two_cols' in missing_rows.csv says which, and how).")
        if combos:
            lines.append(f"{combos} missing row(s) have every value somewhere in the new table, but no new row "
                         f"has them together even when two columns are ignored.")
        if dupes:
            lines.append(f"{dupes} row(s) exist in the new table, but fewer times than in the old one.")
        if closest or to_check:
            n_cols = len(mapping)
            hist = Counter()
            for k in to_check:
                eq = closest.get(k, (0, None))[0]
                hist[n_cols - eq] += 1
            lines += ["", "How close is the closest new row? (columns that differ; 'closest_match_pct' in "
                          "missing_rows.csv)"]
            for d in sorted(hist):
                label = (f"{d} columns differ" if d != 1 else "1 column differs") if d < n_cols else "no similar row found"
                lines.append(f"  {label:<24} {100 * (n_cols - d) / n_cols:>5.0f}% match  {hist[d]:>7} rows")
            by_col = Counter(mapping[i][0] for k in to_check if k in closest
                             for i in range(n_cols) if k[i] != closest[k][1][i])
            if by_col:
                lines.append("  Columns that differ most often from the closest row: "
                             + ", ".join(f"{c} ({n})" for c, n in by_col.most_common(8)))
        if voucher_status:
            st = Counter("whole" if s.startswith("whole") else "none" if s.startswith("voucher in new")
                         else "partial" for s in voucher_status.values())
            whole_v = len({k[V] for k, s in voucher_status.items() if s.startswith("whole")})
            lines += ["", f"By voucher ({mapping[V][0]}; 'voucher_status' in missing_rows.csv):",
                      f"  {st['whole']:>7} missing rows belong to {whole_v} voucher(s) that are not in the new table "
                      f"at all - probably genuinely absent",
                      f"  {st['none']:>7} missing rows belong to vouchers that exist in the new table, but none of "
                      f"their old lines match - look at how the voucher changed",
                      f"  {st['partial']:>7} missing rows are single lines of vouchers whose other lines match"]
        if totals:
            worst = sorted(((b, t) for b, t in totals.items() if t[2]), key=lambda bt: -abs(bt[1][2]))[:10]
            if worst:
                lines += ["", "Largest missing amounts by month" + (" and account" if A is not None else "")
                          + " (all in totals_by_month.csv):"]
                for b, t in worst:
                    lines.append(f"  {' '.join(b):<20}  missing {fmt_decimal(t[2], 2):>14}   old total "
                                 f"{fmt_decimal(t[0], 2):>14}   new total {fmt_decimal(t[1], 2):>14}   "
                                 f"difference {fmt_decimal(t[1] - t[0], 2):>14}")
                lines.append("  ('missing' = amount on old rows not found as-is; a FORMAT row counts here too, "
                             "so compare with 'difference' between the totals)")
                if A is not None and any(i == A for i, _ in recodes):
                    lines.append("  (renamed accounts from recodes.csv are totalled under their new number)")
    excel_err_rows = sum(n for k, n in old_count.items() if any(v.casefold() in EXCEL_ERRORS for v in k))
    dup_rows = sum(n - 1 for n in old_count.values() if n > 1)
    # Progress since the previous run in this results folder.
    previous = None
    try:
        prev = json.loads((out / "compare_results.json").read_text(encoding="utf-8"))
        previous = {"old_rows_missing": prev["old_rows_missing"], "categories": prev.get("categories", {}),
                    "columns_compared": prev.get("columns_compared")}
    except (OSError, ValueError, KeyError):
        pass
    if previous is not None:
        now_cats = Counter(c for c, _, _ in diag.values() if c) if diag else Counter()
        change = missing_rows - previous["old_rows_missing"]
        at = next(i for i, ln in enumerate(lines) if ln.startswith("Old rows NOT in new")) + 1
        lines.insert(at,
                     f"Previous run:          {previous['old_rows_missing']} missing "
                     + ("(no change)" if change == 0 else f"({-change} fewer now)" if change < 0
                        else f"({change} more now)")
                     + ("; different columns compared" if previous["columns_compared"] != len(mapping) else ""))
        moved = [f"{c} {previous['categories'].get(c, 0)} -> {now_cats.get(c, 0)}"
                 for c in ("FORMAT", "CHECK", "REAL", "SCOPE", "EXPLAINED")
                 if previous["categories"].get(c, 0) != now_cats.get(c, 0)]
        if moved:
            lines.insert(at + 1, f"                       by category: {', '.join(moved)}")
    if extra_rows:
        lines += ["", f"The other way round: {extra_rows} new row(s) are not in the old table "
                      f"(extra_in_new.csv)" + (f", counting only the old table's period {lo} - {hi}"
                                               + (" and accounts" if A is not None else "") if D is not None else "")
                  + "."]
        if more_copies:
            lines.append(f"  {more_copies} of them are extra copies of a row that is in the old table fewer times. "
                         f"Check for rows loaded twice into BigQuery.")
        if paired:
            lines.append(f"  {paired} of them are the closest match of a missing old row (1-2 columns differ), so "
                         f"probably the same posting, changed; they are explained in missing_rows.csv.")
        rest = extra_rows - more_copies - paired
        if rest:
            lines.append(f"  {rest} have no counterpart in the old table. Postings made after the Excel copy was "
                         f"taken land here; that is expected.")
        if rest > 0.2 * len(old_keys):
            lines.append("  That is a lot compared with the old table: the new table probably covers more (other "
                         "companies, departments or postings the Excel report filters out). Narrow it ("
                         + ("with the BigQuery filter" if not args.new else "when you export the CSV")
                         + ") to make this list useful.")
    if excel_err_rows or dup_rows:
        lines.append("")
    if excel_err_rows:
        lines.append(f"Excel: {excel_err_rows} old row(s) contain an error value like #N/A or #REF! "
                     f"(difference_type 'excel_error').")
    if dup_rows:
        lines.append(f"Excel: {dup_rows} old row(s) are exact copies of another old row on the compared columns. "
                     f"Real repeated postings are fine; copy-paste leftovers are not.")
    if diag:
        lines.append("\nTip: if one column dominates with a FORMAT difference, fix it (or exclude that column) "
                     "and compare again to see what is left.")
    lines += next_steps(diag, col_kinds, voucher_status, bool(totals), mapping_warnings)
    report = "\n".join(lines)
    write_guide(out)
    (out / "summary.txt").write_text(report + "\n", encoding="utf-8")
    if not args.no_excel:
        t0 = time.time()
        print("Writing results.xlsx ...", file=sys.stderr)
        write_excel_report(out / "results.xlsx", report, [
            ("Missing rows", out / "missing_rows.csv", 100000), ("Totals by month", out / "totals_by_month.csv", 100000),
            ("Renamed values", out / "recodes.csv", 100000), ("Extra in new", out / "extra_in_new.csv", 20000)])
        print(f"  ({time.time() - t0:.0f} s)", file=sys.stderr)
    print("\n" + report)
    print(f"\nWrote {out / 'summary.txt'}, {out / 'HOW_TO_READ_THE_RESULTS.txt'} and {out / 'missing_rows.csv'}"
          + (f" and {out / 'totals_by_month.csv'}" if totals else ""))

    # The same result for other programs: results.json in the results folder (and stdout with --json).
    n_cols = len(mapping)
    closeness = Counter(n_cols - closest.get(k, (0, None))[0] for k in to_check) if (closest or to_check) else {}
    result = {
        "command": "compare",
        "subset": missing_rows == 0,
        "old_rows": len(old_keys), "new_rows": len(new_keys), "columns_compared": n_cols,
        "old_rows_found": matched_rows, "old_rows_missing": missing_rows,
        "distinct_rows_absent": len(fully_missing), "rows_with_fewer_duplicates_in_new": len(short),
        "mapping": [{"old_column": str(o), "new_column": str(n), "type": t}
                    for (o, n), t in zip(mapping, pair_types)],
        "roles": {r: (str(mapping[i][0]) if i is not None else None) for r, i in roles.items()},
        "categories": dict(Counter(c for c, _, _ in diag.values() if c)) if diag else {},
        "not_diagnosed": sum(1 for c, kd, _ in diag.values() if not c and not kd),
        "legitimate_reasons": [{"kind": k, "rows": n, "text": KINDS[k][0], "hint": KINDS[k][2]}
                               for k, n in legit.most_common()],
        "recodes": [{"column": str(mapping[i][0]), "old_value": o, "new_value": nv, "rows": cnt}
                    for (i, o), (nv, cnt) in sorted(recodes.items(), key=lambda x: -x[1][1])],
        "by_column_and_kind": [{"column": str(oc), "kind": kind, "text": describe_kind(kind),
                                "category": KINDS[kind][1].upper() if kind in KINDS else None, "rows": cnt,
                                "example_old": col_example[(oc, kind)][0],
                                "example_new": col_example[(oc, kind)][1] or None}
                               for (oc, kind), cnt in col_kinds.most_common()],
        "closest_match": [{"columns_differing": d, "match_pct": round(100 * (n_cols - d) / n_cols, 1),
                           "rows": closeness[d]} for d in sorted(closeness)],
        "voucher_status": dict(Counter("whole_voucher_missing" if s.startswith("whole")
                                       else "voucher_in_new_no_line_matches" if s.startswith("voucher in new")
                                       else "some_lines_missing" for s in voucher_status.values())),
        "excel_error_rows": excel_err_rows, "excel_duplicate_rows": dup_rows,
        "extra_in_new_rows": extra_rows, "extra_copies_in_new": more_copies,
        "extra_in_new_paired_with_missing": paired,
        "previous_run": previous,
        "mapping_warnings": mapping_warnings,
        "next_steps": [ln.strip() for ln in next_steps(diag, col_kinds, voucher_status, bool(totals), mapping_warnings)
                       if ln.strip() and not ln.strip().startswith("NEXT STEPS")],
        "files": {name: str((out / name).resolve()) for name in
                  ("summary.txt", "results.xlsx", "missing_rows.csv", "totals_by_month.csv", "recodes.csv", "extra_in_new.csv",
                   "HOW_TO_READ_THE_RESULTS.txt", "compare_results.json")
                  if name == "compare_results.json" or (out / name).exists()},
    }
    (out / "compare_results.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result


PLAIN_NUMBER_RE = re.compile(r"^-?(0|[1-9]\d{0,14})(\.\d+)?$")


def write_excel_report(path, summary, sheets):
    """All results in one Excel workbook (a new file in the results folder): the summary, then one sheet per
    CSV with a filter on the header row. Plain numbers become numbers, so they can be summed and sorted."""
    try:
        import openpyxl
        from openpyxl.styles import Font
        from openpyxl.utils import get_column_letter
    except ImportError:
        return
    try:
        wb = openpyxl.Workbook(write_only=True)
        ws = wb.create_sheet("Summary")
        ws.column_dimensions["A"].width = 140
        for ln in summary.splitlines():
            ws.append([ln])
        bold = Font(bold=True)
        for title, csv_path, row_cap in sheets:
            if not Path(csv_path).exists():
                continue
            with open(csv_path, encoding="utf-8-sig", newline="") as f:
                rows = csv.reader(f, delimiter=";")
                header = next(rows, None)
                if not header:
                    continue
                ws = wb.create_sheet(title)
                ws.freeze_panes = "A2"
                widths = [min(max(len(h), 10), 45) for h in header]
                for i, w in enumerate(widths, 1):
                    ws.column_dimensions[get_column_letter(i)].width = w + 2
                cells = []
                for h in header:
                    c = openpyxl.cell.WriteOnlyCell(ws, value=h)
                    c.font = bold
                    cells.append(c)
                ws.append(cells)
                n = 0
                for n, r in enumerate(rows, 1):
                    if n > row_cap:
                        ws.append([f"(only the first {row_cap} rows; all rows are in {Path(csv_path).name})"])
                        break
                    ws.append([float(v) if PLAIN_NUMBER_RE.match(v) else v for v in r])
                ws.auto_filter.ref = f"A1:{get_column_letter(len(header))}{min(n, row_cap) + 1}"
        wb.save(path)
    except Exception as e:  # the CSV files are the real output; a failed workbook must not stop the run
        print(f"warning: could not write {path.name}: {e}", file=sys.stderr)


def cmd_list_columns(args):
    old = _read_file_cols(args.old, args.old_sheet, args.old_header_row, args.old_table, args.old_start_row,
                          args.old_columns)
    original = [str(c) for c in old.columns]
    if args.old_names:
        old = apply_names(old, args.old_names)
    samples = []
    for orig, c in zip(original, old.columns):
        sample = [str(v) for v in old[c].dropna().head(3).tolist()]
        samples.append(" | ".join(sample)[:80])
        label = f"{orig} ({col_letter(int(orig))})" if orig.isdigit() else orig
        if c != orig:
            label += f" = {c}"
        print(f"  {label}: {samples[-1]}")
    print(f"Loaded {len(old)} rows and {len(old.columns)} columns.")
    if args.save:
        old.to_pickle(args.save)
        print(f"Saved a copy for fast reuse: {args.save}")
    print("ROWS:" + str(len(old)))
    print("COLUMNS_JSON:" + json.dumps(list(old.columns), ensure_ascii=False))
    print("ORIGINAL_JSON:" + json.dumps(original, ensure_ascii=False))
    print("SAMPLES_JSON:" + json.dumps(samples, ensure_ascii=False))
    return {"command": "list-columns", "rows": len(old),
            "columns": [{"number": orig, "name": str(c), "sample": smp}
                        for orig, c, smp in zip(original, old.columns, samples)],
            "saved": str(Path(args.save).resolve()) if args.save else None}


def cmd_preview(args):
    """Print column names and sample values of both tables as JSON, for the mapping window."""
    old = read_file(args.old, args.old_sheet, args.old_header_row, args.old_table, args.old_start_row, args.old_columns,
                    args.old_names)
    new = load_new(args, limit=args.sample if (args.new_bq and not args.new) else None)

    def first_values(series, n, distinct):
        """The first n non-empty values of a column (only different ones if distinct)."""
        vals = []
        for v in series.dropna().tolist():
            nv = norm_value(v, args.decimals, False)
            if nv and not (distinct and nv in vals):
                vals.append(nv)
                if len(vals) == n:
                    break
        return vals

    def samples(df):
        return {str(c): [v[:30] for v in first_values(df[c], 4, True)] for c in df.columns}

    def top(df, n=10):
        """Per column, its first n non-empty values (so the rows do not line up across columns)."""
        cols = {str(c): [v[:60] for v in first_values(df[c], n, False)] for c in df.columns}
        depth = max((len(v) for v in cols.values()), default=0)
        return {"columns": list(cols),
                "rows": [[v[i] if i < len(v) else "" for v in cols.values()] for i in range(depth)]}
    print(f"Old: {len(old)} rows, {len(old.columns)} columns. New: {len(new)} rows"
          f"{' (sample)' if args.new_bq and not args.new else ''}, {len(new.columns)} columns.")
    _, candidates = score_pairs(old, new, max(args.sample, 50000), args.decimals, args.casefold,
                                parse_types(args.new_types))
    scores = {}
    for score, oc, nc, name_sim, overlap in candidates:
        scores.setdefault(str(oc), {})[str(nc)] = [round(score, 3), round(overlap, 3), round(name_sim, 3)]
    result = {"old": samples(old), "new": samples(new), "scores": scores, "old_top": top(old), "new_top": top(new)}
    print("PREVIEW_JSON:" + json.dumps(result, ensure_ascii=False))
    return {"command": "preview", "old_rows": len(old), "new_rows": len(new), **result}


def cmd_check_values(args):
    """For each mapped pair: are the old column's non-empty values found in the new column?"""
    mapping = read_mapping(args.mapping)
    types = parse_types(args.new_types)
    old = read_file(args.old, args.old_sheet, args.old_header_row, args.old_table, args.old_start_row, args.old_columns,
                    args.old_names)
    missing_old = [o for o, _ in mapping if o not in old.columns]
    if missing_old:
        sys.exit(f"Columns in mapping not found in old file: {missing_old}. If you renamed or changed "
                 f"the old columns, make the mapping again (Suggest automatically or Edit by hand...).")
    new = load_new(args, columns=sorted({n for _, n in mapping}))
    missing_new = [n for _, n in mapping if n not in new.columns]
    if missing_new:
        sys.exit(f"Columns in mapping not found in new table: {missing_new}")
    print(f"Checking values of {len(mapping)} column pair(s): old {len(old)} rows vs new {len(new)} rows ...",
          file=sys.stderr)

    results, new_sets, not_in = [], {}, []
    excel_rows = bool(old.attrs.get("excel_rows"))
    for oc, nc in mapping:
        if nc not in new_sets:
            new_sets[nc] = set(norm_column(new[nc], args.decimals, args.casefold, types.get(nc)))
            new_sets[nc].discard("")
        normed = norm_column(old[oc], args.decimals, args.casefold, types.get(nc))
        counts = Counter(normed)
        counts.pop("", None)
        absent = [(v, k) for v, k in counts.items() if v not in new_sets[nc]]
        if absent:
            # Where each missing value sits in Excel, and how it looked there before normalising.
            wanted = {v for v, _ in absent}
            seen_rows, raw = {}, {}
            for row_no, v, orig in zip(old.index, normed, old[oc].tolist()):
                if v in wanted:
                    seen_rows.setdefault(v, [])
                    if len(seen_rows[v]) < 5:
                        seen_rows[v].append(row_no)
                    raw.setdefault(v, orig)
            nearest = Nearest(new_sets[nc])
        absent.sort(key=lambda x: -x[1])
        rows = sum(counts.values())
        rows_absent = sum(k for _, k in absent)
        kinds, kind_example = Counter(), {}
        if absent:
            index = FormatIndex(new_sets[nc], [v for v, _ in absent])
            # A date column: a date in a month the new table has no rows for is a scope question, not a format one.
            # (also a period column like 202507)
            months = None
            if len(new_sets[nc]) <= 20000 and \
                    sum(1 for v, _ in absent[:200] if _month_of(v)) > 0.5 * min(len(absent), 200):
                # (a date or period column has at most a few thousand values; voucher numbers have far more)
                new_months = [_month_of(x) for x in new_sets[nc]]
                if sum(1 for ym in new_months if ym) > 0.5 * len(new_sets[nc]):
                    months = {ym for ym in new_months if ym}
            for v, k in absent:
                ym = _month_of(v) if months is not None else None
                if v.casefold() in EXCEL_ERRORS:
                    kind, nv = "excel_error", ""
                elif ym and ym not in months:
                    kind, nv = "no_month", ""
                else:
                    kind, nv = index.explain(v)
                kinds[kind] += k
                kind_example.setdefault(kind, (v, nv))
                not_in.append([oc, nc, v, cell_text(raw.get(v)), k,
                               ", ".join(str(r) for r in seen_rows.get(v, [])) if excel_rows else "",
                               KINDS[kind][1].upper(), KINDS[kind][0], nv or nearest.find(v)])
        causes = [{"kind": kind, "rows": cnt, "nature": KINDS[kind][1], "text": KINDS[kind][0],
                   "hint": KINDS[kind][2], "old": kind_example[kind][0][:40], "new": kind_example[kind][1][:40]}
                  for kind, cnt in kinds.most_common()]
        results.append({"old": oc, "new": nc, "distinct": len(counts), "found": len(counts) - len(absent),
                        "rows": rows, "rows_not_found": rows_absent,
                        "pct": 100.0 * (len(counts) - len(absent)) / len(counts) if counts else None,
                        "examples": [v[:40] for v, _ in absent[:5]],
                        "new_examples": sorted(new_sets[nc])[:3] if absent else [], "causes": causes})

    out = Path(args.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    write_guide(out)
    with open(out / "value_check.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["old_column", "new_column", "old_distinct_values", "found_in_new", "percent_found",
                    "old_rows_with_value", "old_rows_not_found", "examples_not_found", "likely_causes"])
        for r in results:
            w.writerow([r["old"], r["new"], r["distinct"], r["found"],
                        "" if r["pct"] is None else f"{r['pct']:.1f}", r["rows"], r["rows_not_found"],
                        " | ".join(r["examples"]),
                        " | ".join(f"{c['rows']} rows: {c['text']} (e.g. '{c['old']}' vs '{c['new']}')" if c["new"]
                                   else f"{c['rows']} rows: {c['text']}" for c in r["causes"])])

    with open(out / "values_not_in_new.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["old_column", "new_column", "value", "value_in_excel", "rows_in_old",
                    "excel_rows" + (" (first 5)" if excel_rows else ""), "category", "likely_cause",
                    "closest_new_value"])
        w.writerows(not_in)

    width = max(len(f"{r['old']} -> {r['new']}") for r in results)
    print("\nNon-empty old values found in the matched new column:\n")
    for r in results:
        pair = f"{r['old']} -> {r['new']}"
        if r["pct"] is None:
            print(f"  {pair:<{width}}  (old column is empty)")
            continue
        print(f"  {pair:<{width}}  {r['found']}/{r['distinct']} distinct values ({r['pct']:.1f}%)"
              + (f"; not found e.g. {', '.join(repr(e) for e in r['examples'][:3])}" if r["examples"] else ""))
        for c in r["causes"]:
            eg = f"   e.g. old '{c['old']}' vs new '{c['new']}'" if c["new"] else ""
            print(f"  {'':<{width}}    {c['rows']:>7} rows: {describe_kind(c['kind'])}{eg}")
    print(f"\nWrote {out / 'value_check.csv'} and every value not found, one per line, to "
          f"{out / 'values_not_in_new.csv'}.\nA low percentage usually means a wrong match or a format "
          f"difference (e.g. text vs number, date format, decimals).")
    print("NOTIN_FILE:" + str(out / "values_not_in_new.csv"))
    print("CHECK_JSON:" + json.dumps(results, ensure_ascii=False))
    result = {"command": "check-values",
              "all_found": all(r["pct"] is None or r["found"] == r["distinct"] for r in results),
              "columns": results,
              "files": {name: str((out / name).resolve()) for name in
                        ("value_check.csv", "values_not_in_new.csv", "HOW_TO_READ_THE_RESULTS.txt", "value_check.json")}}
    (out / "value_check.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result


# ------------------------------------------------------------------- main ---

def main():
    # Python mixes up the order of sets differently in every run, which made a few diagnoses flip between
    # runs. A fixed hash seed makes the same input give the same output.
    if os.environ.get("PYTHONHASHSEED") != "0" and os.name == "posix":
        os.execve(sys.executable, [sys.executable] + sys.argv, dict(os.environ, PYTHONHASHSEED="0"))
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    def common(p):
        p.add_argument("--old", help="old table (.xlsx/.csv/.pkl); required, on the command line or in --config")
        p.add_argument("--old-sheet")
        p.add_argument("--old-start-row", type=int, help="Excel row where the data starts (default 2); "
                       "columns are then numbered 1, 2, 3 ... instead of named")
        p.add_argument("--old-columns", help="old columns to use by number, e.g. 1-6,8 (default: all)")
        p.add_argument("--old-names", help="names for the old columns, in order, e.g. \"Konto;Tekst;Beløp\" "
                       "(separated by ; or ,; a blank keeps the number)")
        p.add_argument("--old-header-row", type=int, help="instead of numbering, read column names from this row")
        p.add_argument("--old-table", help="name of the Excel table to read, if the sheet has several")
        p.add_argument("--new", help="new table as a file (.xlsx/.csv/.pkl)")
        p.add_argument("--new-sheet")
        p.add_argument("--new-bq", help="BigQuery table, e.g. myproject.finance.Hovedbok")
        p.add_argument("--new-sql", help="file with a BigQuery SELECT to run instead of a whole table")
        p.add_argument("--bq-project", help="Google Cloud project to run the query in")
        p.add_argument("--new-where", help="SQL condition to limit the BigQuery rows, "
                       "e.g. \"bilagsdato BETWEEN '2025-01-01' AND '2025-12-31'\"")
        p.add_argument("--cache", help="save/reuse the BigQuery result in this .pkl file")
        p.add_argument("--decimals", type=int, default=2, help="round numbers to this many decimals (default 2)")
        p.add_argument("--casefold", action="store_true", help="ignore upper/lower case in text")
        p.add_argument("--new-types", help="data types of new (BigQuery) columns, e.g. "
                       "\"konto=STRING; belop=NUMERIC; bilagsdato=DATE\". Types: STRING, INTEGER, NUMERIC, DATE, "
                       "TIMESTAMP, BOOL. The Excel column mapped to it is read the same way.")
        p.add_argument("--columns-file", help="text file with the old columns to compare, one per line (default: all)")
        p.add_argument("--skip-columns-file", help="text file with old columns to leave out, one per line")

    s = sub.add_parser("suggest", help="propose a column mapping")
    common(s)
    s.add_argument("--out", default="mapping.csv")
    s.add_argument("--sample", type=int, default=50000, help="rows to sample for value overlap (default 50000)")
    s.set_defaults(func=cmd_suggest)

    c = sub.add_parser("compare", help="check that old rows are contained in new")
    common(c)
    c.add_argument("--mapping", default="mapping.csv")
    c.add_argument("--out-dir", default="hovedbok_compare_results")
    for role in ("date", "voucher", "account", "amount"):
        c.add_argument(f"--{role}-col", help=f"old column that holds the {role} (default: guessed)")
    c.add_argument("--diagnose-limit", type=int, default=50000,
                   help="how many missing rows to diagnose for near matches (default 50000)")
    c.add_argument("--no-excel", action="store_true", help="don't write results.xlsx (the CSV files are always written)")
    c.set_defaults(func=cmd_compare)

    cv = sub.add_parser("check-values", help="check that each old column's values exist in its matched new column")
    common(cv)
    cv.add_argument("--mapping", default="mapping.csv")
    cv.add_argument("--out-dir", default="hovedbok_compare_results")
    cv.set_defaults(func=cmd_check_values)

    pv = sub.add_parser("preview", help="print columns and sample values of both tables (for manual mapping)")
    common(pv)
    pv.add_argument("--sample", type=int, default=20000,
                    help="rows fetched from BigQuery for the preview and match scores (default 20000)")
    pv.set_defaults(func=cmd_preview)

    lc = sub.add_parser("list-columns", help="print the old file's column names")
    lc.add_argument("--old")
    lc.add_argument("--old-sheet")
    lc.add_argument("--old-header-row", type=int)
    lc.add_argument("--old-start-row", type=int)
    lc.add_argument("--old-columns")
    lc.add_argument("--old-names")
    lc.add_argument("--old-table")
    lc.add_argument("--save", help="also save the loaded table as a .pkl file; pass it as --old later to skip re-reading Excel")
    lc.set_defaults(func=cmd_list_columns, new=None, new_bq=None, new_sql=None)

    k = sub.add_parser("kinds", help="list every difference type the diagnostics can report (for programs/AI)")
    k.set_defaults(func=cmd_kinds)

    rn = sub.add_parser("run", help="everything in one go: suggest a mapping if there is none, then check-values "
                        "and compare (reads each table once)")
    common(rn)
    rn.add_argument("--mapping", default="mapping.csv", help="mapping to use; made with 'suggest' if it doesn't exist")
    rn.add_argument("--suggest", action="store_true", help="make a new mapping even if the file exists")
    rn.add_argument("--sample", type=int, default=50000, help="rows to sample when suggesting (default 50000)")
    rn.add_argument("--out-dir", default="hovedbok_compare_results")
    for role in ("date", "voucher", "account", "amount"):
        rn.add_argument(f"--{role}-col", help=f"old column that holds the {role} (default: guessed)")
    rn.add_argument("--diagnose-limit", type=int, default=50000)
    rn.add_argument("--no-excel", action="store_true", help="don't write results.xlsx")
    rn.set_defaults(func=cmd_run)

    st = sub.add_parser("selftest", help="make a small test table with known differences and check that the "
                        "comparison finds them (checks the installation; touches none of your files)")
    st.add_argument("--keep", help="keep the test files in this folder instead of deleting them")
    st.set_defaults(func=cmd_selftest)

    subparsers = {"suggest": s, "compare": c, "check-values": cv, "preview": pv, "list-columns": lc, "kinds": k,
                  "selftest": st, "run": rn}
    for p in subparsers.values():
        p.add_argument("--config", help="JSON file with options, e.g. {\"old\": \"Hovedbok.xlsm\", "
                       "\"old-start-row\": 16}; options on the command line win")
        p.add_argument("--json", action="store_true", help="print one JSON result on stdout; the usual text "
                       "goes to stderr. Errors become {\"ok\": false, \"error\": ...}")
    for p in (c, cv, rn):
        p.add_argument("--fail-on-missing", action="store_true",
                       help="exit with code 2 if any old row (compare) or value (check-values) is missing")

    argv = sys.argv[1:]
    pre, _ = ap.parse_known_args(argv)
    if pre.config:
        apply_config(subparsers[pre.cmd], pre.config)
    args = ap.parse_args(argv)
    sys.exit(run(args))


def apply_config(parser, path):
    """Use the options in a JSON settings file as defaults for this command."""
    try:
        cfg = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    except (OSError, ValueError) as e:
        sys.exit(f"Can't read the settings file {path}: {e}")
    if not isinstance(cfg, dict):
        sys.exit(f"The settings file {path} must hold one JSON object, like {{\"old\": \"Hovedbok.xlsm\"}}")
    known = {a.dest: a for a in parser._actions}
    defaults, unknown = {}, []
    for key, value in cfg.items():
        dest = key.lstrip("-").replace("-", "_")
        if dest not in known or dest in ("help", "config", "func"):
            unknown.append(key)
            continue
        if isinstance(value, (list, tuple)):  # e.g. "old-names": ["Konto", "Tekst"]
            value = ";".join(str(v) for v in value)
        elif isinstance(value, dict):  # e.g. "new-types": {"belop": "NUMERIC"}
            value = "; ".join(f"{c}={t}" for c, t in value.items())
        defaults[dest] = value
    if unknown:
        print(f"warning: settings not used by this command: {', '.join(unknown)}", file=sys.stderr)
    parser.set_defaults(**defaults)


def run(args):
    """Run the chosen command. Returns the exit code: 0 ok, 1 error, 2 missing rows/values (--fail-on-missing)."""
    as_json = getattr(args, "json", False)
    stdout = sys.stdout
    if as_json:
        try:
            stdout.reconfigure(encoding="utf-8")  # JSON is UTF-8, also on a Windows console
        except (AttributeError, ValueError):
            pass
        sys.stdout = sys.stderr  # keep stdout for the JSON alone
    try:
        if args.func not in (cmd_kinds, cmd_selftest) and not args.old:
            sys.exit("Give the old table with --old FILE (or \"old\" in the --config file).")
        if getattr(args, "new", None) and (getattr(args, "new_bq", None) or getattr(args, "new_sql", None)):
            sys.exit("Use either --new or --new-bq/--new-sql, not both.")
        result = args.func(args) or {}
        code = 0
        if result.get("passed") is False:
            code = 2
        if getattr(args, "fail_on_missing", False) and (result.get("subset") is False
                                                        or result.get("all_found") is False):
            code = 2
    except SystemExit as e:
        if e.code in (None, 0):
            result, code = {}, 0
        else:
            msg = e.code if isinstance(e.code, str) else f"exit code {e.code}"
            if not as_json:
                raise
            result, code = {"error": msg}, 1
    except KeyboardInterrupt:
        result, code = {"error": "stopped"}, 1
        if not as_json:
            raise
    except Exception as e:
        if not as_json:
            raise
        import traceback
        traceback.print_exc()
        result, code = {"error": f"{type(e).__name__}: {e}"}, 1
    finally:
        sys.stdout = stdout
    if as_json:
        print(json.dumps({"ok": code != 1, "exit_code": code, **result}, ensure_ascii=False, indent=2,
                         default=str))
    return code


def cmd_run(args):
    """suggest (when needed) + check-values + compare, reading each table only once."""
    global _MEMO
    _MEMO = {}
    result = {"command": "run", "mapping_file": str(Path(args.mapping).resolve())}
    try:
        if args.suggest or not Path(args.mapping).exists():
            print(f"== No mapping yet: suggesting one ({args.mapping}) ==" if not args.suggest
                  else "== Suggesting a new mapping ==", file=sys.stderr)
            args.out = args.mapping
            result["suggest"] = cmd_suggest(args)
            result["mapping_was_suggested"] = True
            print("NOTE: the mapping was made automatically and nobody has checked it yet. If a column is matched "
                  "wrongly, the results below are wrong too.", file=sys.stderr)
        else:
            result["mapping_was_suggested"] = False
        print("\n== Checking values column by column ==")
        result["check_values"] = cmd_check_values(args)
        print("\n== Comparing whole rows ==")
        result["compare"] = cmd_compare(args)
    finally:
        _MEMO = None
    result["subset"] = result["compare"]["subset"]
    result["all_found"] = result["check_values"]["all_found"]
    return result


def make_test_data(folder, seed=7):
    """A small old Hovedbok (Excel, data from row 3) and a new one (CSV, BigQuery style) with planted
    differences. Returns {planted case: (old row number(s) in Excel, expected difference_type)}."""
    import random
    rnd = random.Random(seed)
    accounts = ["1920", "3000", "4000", "6300", "6800", "7140"]
    old_rows, new_rows, planted = [], [], {}
    voucher = 50000

    def line(day, acc, text, amount, vno):
        return [vno, dt.date(2025, 1, 1) + dt.timedelta(days=day), acc, text, round(amount, 2)]
    for _ in range(150):  # ordinary vouchers: two lines that balance, identical in old and new
        voucher += 1
        day, amt = rnd.randrange(360), rnd.uniform(100, 20000)
        acc = rnd.choice([a for a in accounts if a != "6300"])
        text = f"Faktura {rnd.randrange(10000, 99999)}"
        old_rows += [line(day, acc, text, amt, voucher), line(day, "1920", text, -amt, voucher)]
    new_rows = [list(r) for r in old_rows]

    def plant(name, kind, old, new):
        planted[name] = (len(old_rows) + 3, kind)
        old_rows.append(old)
        if new is not None:
            new_rows.append(new)
    voucher += 1
    plant("case", "case", line(10, "4000", "Kontorrekvisita", 512.5, voucher),
          line(10, "4000", "KONTORREKVISITA", 512.5, voucher))
    voucher += 1
    plant("sign", "sign", line(11, "4000", "Kreditnota 77", 300, voucher), line(11, "4000", "Kreditnota 77", -300, voucher))
    voucher += 1
    plant("day_shift", "day_shift", line(40, "6800", "Husleie mars", 18000, voucher),
          line(39, "6800", "Husleie mars", 18000, voucher))
    voucher += 1
    plant("rounding", "rounding", line(50, "7140", "Reise Oslo", 1234.57, voucher),
          line(50, "7140", "Reise Oslo", 1234.56, voucher))
    voucher += 1
    plant("out_of_period", "out_of_period", [voucher, dt.date(2023, 6, 1), "3000", "Gammel post", 999.0], None)
    voucher += 1
    for i, acc in enumerate(("3000", "1920")):  # a whole, balanced voucher missing from new: a REAL difference
        plant(f"missing_voucher_{i}", "missing_voucher",
              line(200, acc, "Salg uten match", 4321.0 * (-1 if i else 1), voucher), None)
    voucher += 1
    for i, amt in enumerate((700, -700)):  # posted and reversed on the same account: legitimately left out
        plant(f"cancelled_{i}", "cancelled_voucher", line(210, "4000", "Feilført", amt, voucher), None)
    for i in range(6):  # account renamed 6300 -> 6310 in the new system
        voucher += 1
        r = line(100 + i, "6300", f"Leie maskin {i}", 1500 + i, voucher)
        plant(f"recode_{i}", "recode", r, r[:2] + ["6310"] + r[3:])
    voucher += 1  # a VAT line is not an opening balance: missing, so REAL
    plant("inngaende_mva", "missing_voucher", line(220, "2710", "Inngående mva", 125.0, voucher), None)
    for i, amt in enumerate((640, -640)):  # posted, then reversed on a voucher of its own; neither is in new
        voucher += 1
        plant(f"reversal_{i}", "reversal", line(230, "7140", "Taxi feil", amt, voucher), None)
    voucher += 1
    plant("excel_error", "excel_error", line(240, "6800", "#N/A", 99.0, voucher), None)
    voucher += 1
    r = line(300, "4000", "Dobbel linje", 250, voucher)
    plant("duplicates", "duplicates", r, r)
    old_rows.append(list(r))  # twice in old, once in new
    new_rows.append(list(new_rows[0]))  # loaded twice into the new table

    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bilagssøk"
    ws.append(["Test Hovedbok (made by selftest)"])
    ws.append(["Bilagsnr", "Bilagsdato", "Konto", "Tekst", "Beløp"])
    for r in old_rows:
        ws.append([r[0], r[1], int(r[2]), r[3], r[4]])
        if r[3] == "#N/A":  # a real Excel error cell, as a broken formula leaves it
            ws.cell(row=ws.max_row, column=4).data_type = "e"
    for c in ws["B"][2:]:
        c.number_format = "dd.mm.yyyy"
    wb.save(folder / "old.xlsx")
    with open(folder / "new.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["bilagsnr", "bilagsdato", "konto", "tekst", "belop", "firma"])
        for r in new_rows:  # BigQuery style: timestamp with UTC offset, an extra column
            w.writerow([r[0], f"{r[1]:%Y-%m-%d} 00:00:00+00:00", r[2], r[3], f"{r[4]:.2f}", "USBL"])
    return planted, len(old_rows), len(new_rows)


def cmd_selftest(args):
    """Check that the program works on this computer: make test data, run it, compare with the known answer."""
    import subprocess
    import tempfile
    tmp = None
    if args.keep:
        folder = Path(args.keep)
        folder.mkdir(parents=True, exist_ok=True)
    else:
        tmp = tempfile.TemporaryDirectory()
        folder = Path(tmp.name)
    try:
        planted, n_old, n_new = make_test_data(folder)
        me = [sys.executable, str(Path(__file__).resolve())]
        src = ["--old", str(folder / "old.xlsx"), "--old-start-row", "3",
               "--old-names", "bilagsnr;bilagsdato;konto;tekst;belop", "--new", str(folder / "new.csv")]
        checks = []

        def step(name, cmd):
            t0 = time.time()
            p = subprocess.run(me + cmd + ["--json"], capture_output=True, text=True, encoding="utf-8",
                               errors="replace", env=dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONHASHSEED="0"))
            try:
                res = json.loads(p.stdout)
            except ValueError:
                res = {"ok": False, "error": (p.stderr or p.stdout)[-2000:]}
            print(f"  {name:<14} {'ok' if res.get('ok') else 'FAILED'} ({time.time() - t0:.1f} s)")
            if not res.get("ok"):
                print("    " + str(res.get("error")).replace("\n", "\n    "))
            return res

        print(f"Test data: {n_old} old rows (Excel), {n_new} new rows (CSV) in {folder}")
        sug = step("suggest", ["suggest"] + src + ["--out", str(folder / "mapping.csv")])
        if sug.get("ok"):
            wrong = [m["old_column"] for m in sug["mapping"] if m["new_column"] != m["old_column"]]
            checks.append(("suggest matched every column", not wrong, f"wrong: {wrong}" if wrong else ""))
        else:
            # Write the mapping by hand so the rest can still be tested.
            (folder / "mapping.csv").write_text("old_column;new_column\n" + "".join(
                f"{c};{c}\n" for c in ("bilagsnr", "bilagsdato", "konto", "tekst", "belop")), encoding="utf-8")
        cv = step("check-values", ["check-values"] + src + ["--mapping", str(folder / "mapping.csv"),
                                                          "--out-dir", str(folder / "res")])
        if cv.get("ok"):
            not_found = {c["old"]: c["examples"] for c in cv["columns"] if c["examples"]}
            checks.append(("check-values: only konto 6300, 2710 and the planted values are missing",
                           set(not_found.get("konto", [])) == {"6300", "2710"} and "bilagsnr" in not_found,
                           f"not found: {not_found}"))
        cmp = step("compare", ["compare"] + src + ["--mapping", str(folder / "mapping.csv"),
                                                 "--out-dir", str(folder / "res"), "--new-types",
                                                 "bilagsdato=DATE_LOCAL"])
        if cmp.get("ok"):
            expected_missing = len(planted)  # every planted row; the duplicate counts once
            checks.append(("rows missing", cmp["old_rows_missing"] == expected_missing,
                           f"expected {expected_missing}, got {cmp['old_rows_missing']}"))
            with open(folder / "res" / "missing_rows.csv", encoding="utf-8-sig") as f:
                found, cats = {}, {}
                for r in csv.DictReader(f, delimiter=";"):
                    for no in r["excel_rows"].replace(" ...", "").split(", "):
                        cats[int(no)] = r["category"]
                    for no in r["excel_rows"].replace(" ...", "").split(", "):
                        found[int(no)] = r["difference_type"]
            for name, (row, kind) in planted.items():
                got = found.get(row, "(not in missing_rows.csv)")
                cat = cats.get(row, "")
                ok = (cat == "REAL") if kind == "missing_voucher" else kind in got.split(",")
                if kind == "missing_voucher":
                    got = f"{cat} {got}"
                checks.append((f"{name} found as {'REAL' if kind == 'missing_voucher' else kind}", ok, "" if ok else f"got '{got}'"))
            checks.append(("a row loaded twice into new is found", cmp.get("extra_copies_in_new") == 1,
                           f"got {cmp.get('extra_copies_in_new')}"))
            checks.append(("totals_by_month.csv written", (folder / "res" / "totals_by_month.csv").exists(), ""))
    finally:
        if tmp:
            tmp.cleanup()
    print("\nChecks:")
    for name, ok, why in checks:
        print(f"  {'ok    ' if ok else 'FAILED'}  {name}" + (f"  ({why})" if why else ""))
    failed = [n for n, ok, _ in checks if not ok]
    print("\nSELFTEST " + ("PASSED: the program works on this computer." if not failed and checks else
                           f"FAILED: {len(failed)} check(s) failed. Send this output to whoever maintains the script."))
    return {"command": "selftest", "passed": bool(checks) and not failed,
            "checks": [{"check": n, "ok": ok, "detail": why} for n, ok, why in checks],
            "kept_in": str(folder.resolve()) if args.keep else None}


def cmd_kinds(args):
    """List the difference types, so a program or AI can interpret difference_type values."""
    kinds = [{"kind": k, "text": t, "category": nat.upper(), "hint": h} for k, (t, nat, h) in KINDS.items()]
    for kd in kinds:
        print(f"  {kd['kind']:<20} {kd['category']:<9} {kd['text']}  ({kd['hint']})")
    return {"command": "kinds", "categories": {
        "FORMAT": "same data, written differently - fix the format, not the data",
        "CHECK": "needs a look (blanks, near-identical text, duplicates, values that exist but not together)",
        "REAL": "a genuine difference: a value differs, or is not in the new table at all",
        "SCOPE": "outside what the new table covers (date range or missing month) - not an error",
        "EXPLAINED": "a likely legitimate reason was found - confirm, don't fix"}, "kinds": kinds}


if __name__ == "__main__":
    main()
