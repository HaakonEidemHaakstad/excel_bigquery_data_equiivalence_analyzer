#!/usr/bin/env python3
"""
bq_download.py - download a BigQuery table or query result to one CSV file,
page by page, with no row limit (the console download stops at a few MB).

Examples:
  python bq_download.py --table usbl-innsikt.unit4_datasett.hovedbok --out hovedbok_bq.csv
  python bq_download.py --table usbl-innsikt.unit4_datasett.hovedbok --where "bilagsdato >= '2025-01-01'" --out hovedbok_2025.csv
  python bq_download.py --sql query.sql --out result.csv

Options:
  --table PROJECT.DATASET.TABLE  table to download
  --columns "a,b,c"              only these columns (default: all)
  --where "CONDITION"            filter, e.g. "bilagsdato >= '2025-01-01'"
  --sql FILE.sql                 run your own SELECT instead of --table
  --project PROJECT              Google Cloud project that runs the query
  --page-size 50000              rows fetched per batch
  --sep ";"                      CSV separator (default ",")

Needs: pip install google-cloud-bigquery pydata-google-auth
Sign-in: the first run opens your browser to log in with your Google account
(no gcloud needed). The login is saved for later runs.
Read-only: only runs a SELECT.
"""

import argparse
import csv
import datetime as dt
import decimal
import sys
import time
from pathlib import Path


def cell(v):
    if v is None:
        return ""
    if isinstance(v, dt.datetime):
        return v.isoformat(sep=" ")
    if isinstance(v, (dt.date, dt.time)):
        return v.isoformat()
    if isinstance(v, decimal.Decimal):
        return format(v, "f")
    if isinstance(v, bytes):
        return v.hex()
    return v


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


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--table")
    ap.add_argument("--columns")
    ap.add_argument("--where")
    ap.add_argument("--sql")
    ap.add_argument("--project")
    ap.add_argument("--out", required=True)
    ap.add_argument("--page-size", type=int, default=50000)
    ap.add_argument("--sep", default=",")
    args = ap.parse_args()

    if bool(args.table) == bool(args.sql):
        sys.exit("Give either --table or --sql.")
    try:
        from google.cloud import bigquery
    except ImportError:
        sys.exit("Missing package. Install with:  python -m pip install google-cloud-bigquery pydata-google-auth")

    if args.sql:
        sql = Path(args.sql).read_text(encoding="utf-8")
    else:
        cols = ", ".join(f"`{c.strip()}`" for c in args.columns.split(",")) if args.columns else "*"
        sql = f"SELECT {cols} FROM `{args.table.strip('`')}`"
        if args.where:
            sql += f" WHERE {args.where}"

    print(f"Running:\n  {sql}\n", file=sys.stderr)
    project = args.project or (args.table.strip('`').split('.')[0] if args.table and args.table.count('.') == 2 else None)
    client = make_bq_client(bigquery, project)
    job = client.query(sql)
    rows = job.result(page_size=args.page_size)
    total = rows.total_rows
    print(f"{total} rows to download.", file=sys.stderr)

    start, written = time.time(), 0
    with open(args.out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=args.sep)
        w.writerow([field.name for field in rows.schema])
        for page in rows.pages:
            for row in page:
                w.writerow([cell(v) for v in row.values()])
                written += 1
            print(f"  {written}/{total} rows  ({time.time() - start:.0f}s)", file=sys.stderr)

    print(f"\nDone: wrote {written} rows to {args.out}", file=sys.stderr)
    if total is not None and written != total:
        print(f"WARNING: expected {total} rows but wrote {written}.", file=sys.stderr)


if __name__ == "__main__":
    main()
