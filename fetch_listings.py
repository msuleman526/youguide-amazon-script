#!/usr/bin/env python3
"""
Amazon SP-API → Google Sheet exporter  (full-account, Reports API)
==================================================================
Pulls EVERY listing in the seller account for the target marketplace and writes
them — with all columns Amazon provides — into a SEPARATE worksheet/tab.

Why the Reports API instead of searchListingsItems:
    The Listings Items `searchListingsItems` operation reports the true total
    (e.g. 984) but returns only a single page and never hands back a pagination
    token for this account, so it cannot retrieve the full catalogue. The
    Merchant Listings report (GET_MERCHANT_LISTINGS_ALL_DATA) returns every SKU
    in one flat file — the same data behind Seller Central's listings reports.

This is read-only against Amazon: it never creates, updates, or deletes a
listing. It is also non-destructive to your sheet — it writes to its own tab
(TARGET_WORKSHEET) and leaves the uploader's `amazon_listings_template` tab,
including any HOLD flags, untouched.

Run it before the uploader to see everything already live on Amazon:
    python fetch_listings.py

Dependencies are the same as the uploader (python-amazon-sp-api, gspread).
"""

import csv
import io
import logging
import time
from typing import Optional

import gspread
from sp_api.api import Reports
from sp_api.base import SellingApiException
from sp_api.base.reportTypes import ReportType

# Reuse every credential / config value already set up in the uploader so there
# is only one place to maintain secrets and IDs.
from upload_listings import (
    SP_API_CREDENTIALS,
    MARKETPLACE,
    MARKETPLACE_ID,
    SPREADSHEET_ID,
    OAUTH_CREDENTIALS_FILE,
    OAUTH_TOKEN_FILE,
)


# =============================================================================
#  CONFIGURATION
# =============================================================================

# The tab this script writes to. Created automatically if it does not exist and
# fully cleared+rewritten on each run. Keep it DIFFERENT from the uploader's
# WORKSHEET_NAME ("amazon_listings_template") so nothing you authored is lost.
TARGET_WORKSHEET = "amazon_current"

# GET_MERCHANT_LISTINGS_ALL_DATA = every listing (active + inactive) with the
# full Seller-Central flat-file column set. Swap for GET_MERCHANT_LISTINGS_DATA
# for active-only, or GET_FLAT_FILE_OPEN_LISTINGS_DATA for the lite open set.
REPORT_TYPE = ReportType.GET_MERCHANT_LISTINGS_ALL_DATA

# Report generation is asynchronous — poll until DONE.
POLL_INTERVAL_SECONDS = 5
POLL_TIMEOUT_SECONDS  = 600   # give up after 10 minutes

# A few network calls can hit transient TLS timeouts — retry them.
HTTP_RETRIES   = 5
HTTP_BACKOFF_S = 3

# Columns hoisted to the front of the sheet for readability; every remaining
# report column follows in Amazon's native order, so nothing is dropped.
PRIORITY_COLUMNS = ["seller-sku", "asin1", "item-name", "price", "quantity", "status"]


# =============================================================================
#  LOGGING
# =============================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)


# =============================================================================
#  HELPERS
# =============================================================================

def with_retry(fn, *args, **kwargs):
    """Call fn, retrying on transient network errors (flaky TLS handshakes)."""
    for attempt in range(1, HTTP_RETRIES + 1):
        try:
            return fn(*args, **kwargs)
        except SellingApiException:
            raise  # real API errors are not retried here
        except Exception as exc:
            if attempt == HTTP_RETRIES:
                raise
            log.warning("Network hiccup (%s) — retry %d/%d", exc.__class__.__name__, attempt, HTTP_RETRIES)
            time.sleep(HTTP_BACKOFF_S)


# =============================================================================
#  GOOGLE SHEETS
# =============================================================================

def get_spreadsheet() -> gspread.Spreadsheet:
    """Open the spreadsheet via the same OAuth user credentials as the uploader."""
    client = gspread.oauth(
        credentials_filename=OAUTH_CREDENTIALS_FILE,
        authorized_user_filename=OAUTH_TOKEN_FILE,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )
    return client.open_by_key(SPREADSHEET_ID)


def get_target_worksheet(spreadsheet: gspread.Spreadsheet, n_rows: int, n_cols: int) -> gspread.Worksheet:
    """Return the export tab, creating it if needed, and clear any old contents."""
    rows = max(n_rows + 10, 100)
    try:
        ws = spreadsheet.worksheet(TARGET_WORKSHEET)
        ws.clear()
        if ws.col_count < n_cols:
            ws.add_cols(n_cols - ws.col_count)
        if ws.row_count < rows:
            ws.add_rows(rows - ws.row_count)
        log.info("Reusing existing tab %r (cleared)", TARGET_WORKSHEET)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=TARGET_WORKSHEET, rows=rows, cols=n_cols)
        log.info("Created new tab %r", TARGET_WORKSHEET)
    return ws


# =============================================================================
#  REPORTS API
# =============================================================================

def request_report(client: Reports) -> str:
    """Create the merchant-listings report and return its reportId."""
    resp = with_retry(
        client.create_report,
        reportType=REPORT_TYPE if isinstance(REPORT_TYPE, str) else REPORT_TYPE.value,
        marketplaceIds=[MARKETPLACE_ID],
    )
    report_id = (resp.payload or {}).get("reportId")
    if not report_id:
        raise RuntimeError(f"create_report returned no reportId: {resp.payload}")
    log.info("Report requested — reportId=%s", report_id)
    return report_id


def wait_for_report(client: Reports, report_id: str) -> str:
    """Poll until the report is DONE; return its reportDocumentId."""
    deadline = POLL_TIMEOUT_SECONDS
    waited = 0
    while True:
        resp    = with_retry(client.get_report, report_id)
        payload = resp.payload or {}
        status  = payload.get("processingStatus", "UNKNOWN")
        log.info("Report %s status: %s (waited %ds)", report_id, status, waited)

        if status == "DONE":
            doc_id = payload.get("reportDocumentId")
            if not doc_id:
                raise RuntimeError(f"Report DONE but no reportDocumentId: {payload}")
            return doc_id
        if status in ("CANCELLED", "FATAL"):
            raise RuntimeError(f"Report {report_id} ended with status {status}: {payload}")

        if waited >= deadline:
            raise TimeoutError(f"Report {report_id} not done after {deadline}s (last status {status})")
        time.sleep(POLL_INTERVAL_SECONDS)
        waited += POLL_INTERVAL_SECONDS


def download_report(client: Reports, document_id: str) -> str:
    """Download + decompress the report document; return its TSV text."""
    resp = with_retry(client.get_report_document, document_id, download=True)
    text = (resp.payload or {}).get("document")
    if text is None:
        raise RuntimeError(f"No document content for {document_id}: {resp.payload}")
    return text


def parse_report(text: str) -> tuple[list, list]:
    """Parse the tab-separated report into (header, rows)."""
    reader = csv.reader(io.StringIO(text), delimiter="\t")
    all_rows = [r for r in reader if r]
    if not all_rows:
        return [], []
    header, rows = all_rows[0], all_rows[1:]
    return header, rows


def reorder_columns(header: list, rows: list) -> tuple[list, list]:
    """Move PRIORITY_COLUMNS to the front; keep all others in native order."""
    present_priority = [c for c in PRIORITY_COLUMNS if c in header]
    rest = [c for c in header if c not in present_priority]
    new_header = present_priority + rest
    index = {c: header.index(c) for c in header}

    def reshape(row: list) -> list:
        # pad short rows so column alignment holds even if Amazon omits trailing fields
        padded = row + [""] * (len(header) - len(row))
        return [padded[index[c]] for c in new_header]

    return new_header, [reshape(r) for r in rows]


# =============================================================================
#  MAIN
# =============================================================================

def run() -> None:
    log.info("=== Amazon → Google Sheet listing export (Reports API) ===")
    log.info("Marketplace: %s (%s) | Report: %s",
             MARKETPLACE.name, MARKETPLACE_ID,
             REPORT_TYPE if isinstance(REPORT_TYPE, str) else REPORT_TYPE.value)

    client = Reports(credentials=SP_API_CREDENTIALS, marketplace=MARKETPLACE)

    report_id   = request_report(client)
    document_id = wait_for_report(client, report_id)
    log.info("Report ready — downloading document %s", document_id)
    text = download_report(client, document_id)

    header, rows = parse_report(text)
    if not rows:
        log.warning("Report contained no listing rows — nothing written.")
        return
    log.info("Parsed %d listings (%d columns)", len(rows), len(header))

    header, rows = reorder_columns(header, rows)
    rows.sort(key=lambda r: r[0])  # sort by seller-sku

    log.info("Writing %d rows to tab %r…", len(rows), TARGET_WORKSHEET)
    spreadsheet = get_spreadsheet()
    worksheet   = get_target_worksheet(spreadsheet, n_rows=len(rows), n_cols=len(header))
    worksheet.update(values=[header] + rows, range_name="A1")

    log.info("Done. %d listings exported to %r.", len(rows), TARGET_WORKSHEET)


if __name__ == "__main__":
    try:
        run()
    except SellingApiException as exc:
        log.error("SP-API error: %s", exc)
        log.error("If this is a 403, the app/token is missing the report role "
                  "(run diagnose.py to confirm auth).")
