#!/usr/bin/env python3
"""
Amazon SP-API ⇄ Google Sheet two-way sync — YouGuide eSIM / SIM Cards
=====================================================================
Keeps a Google Sheet and your Amazon UK listings in sync, both directions,
on a continuous activity-based poll (5–10 min):

  Sheet → Amazon
    * Edit ANY content column → the listing is updated on Amazon.
    * Action column (O) set to DELETE/X/TRUE → the listing is removed.
    * A previously-DELETED row whose action is cleared / set FALSE → re-created.

  Amazon → Sheet  (reverse sync)
    * If a listing is changed directly on Amazon (e.g. in Seller Central) since
      the last sync, the change is written back into the sheet.
    * If a listing is removed on Amazon, the row is marked "DELETED (on Amazon)".

Conflict rule: the SHEET is the primary source of truth. If only Amazon changed
→ pull into the sheet. If only the sheet changed → push to Amazon. If BOTH
changed since the last sync → the sheet wins and the Amazon value + its
timestamp are recorded in the CONFLICT column for manual review.

Install dependencies:
    pip install python-amazon-sp-api gspread google-auth

Usage:
    python upload_listings.py
"""

import hashlib
import json
import logging
import os
import time
from datetime import datetime, timezone
from typing import Optional

import gspread
from sp_api.api import ListingsItems
from sp_api.base import Marketplaces, SellingApiException


# =============================================================================
#  SECRETS LOADING (.env)
#  Secrets live in a .env file next to this script — NOT in source — so the
#  code can be shared/committed without leaking the refresh token or LWA secret.
#  A tiny built-in parser is used so python-dotenv is not a hard dependency.
# =============================================================================

def load_dotenv(path: str = ".env") -> None:
    """Read KEY=VALUE lines from .env into os.environ (existing vars win)."""
    here = os.path.join(os.path.dirname(os.path.abspath(__file__)), path)
    if not os.path.exists(here):
        return
    with open(here, encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            os.environ.setdefault(key.strip(), value.strip())


load_dotenv()


def _require(name: str) -> str:
    """Fetch a required secret from the environment, warning if it is missing."""
    val = os.getenv(name, "")
    if not val:
        log_missing = logging.getLogger(__name__)
        log_missing.warning("Missing %s — set it in .env before running for real", name)
    return val


# =============================================================================
#  CREDENTIALS & CONFIGURATION
# =============================================================================

# Amazon SP-API OAuth credentials  (from Seller Central → Apps & Services →
# Develop Apps → your app → View credentials). Values come from .env.
SP_API_CREDENTIALS = {
    "refresh_token":     _require("SP_API_REFRESH_TOKEN"),
    "lwa_app_id":        _require("LWA_APP_ID"),
    "lwa_client_secret": _require("LWA_CLIENT_SECRET"),
}

# Your Amazon Seller account ID  (Seller Central → Settings → Account Info)
SELLER_ID = _require("SELLER_ID")

# Target marketplace — existing YouGuide listings are on Amazon UK
MARKETPLACE = Marketplaces.UK

# Google Sheets — OAuth user auth (no service account key needed).
# Create an OAuth 2.0 Client ID of type "Desktop app" in Google Cloud Console
# (APIs & Services → Credentials → Create Credentials → OAuth client ID) and
# download its JSON to OAUTH_CREDENTIALS_FILE. On first run a browser opens for
# you to log in; the resulting token is cached in OAUTH_TOKEN_FILE for reuse.
OAUTH_CREDENTIALS_FILE = "oauth_client.json"       # downloaded OAuth client secrets
OAUTH_TOKEN_FILE       = "authorized_user.json"    # auto-created after first login
SPREADSHEET_ID         = _require("SPREADSHEET_ID")  # from sheet URL (set in .env)
WORKSHEET_NAME         = "amazon_listings_template"  # Tab / worksheet name

# Safety switch — when True, NOTHING is written to Amazon or the sheet; the
# script only logs the create/update/delete/pull decisions it WOULD make (live
# READS still happen, since they're harmless). Start with True to confirm the
# reverse-sync behaviour, then set to False to go live.
DRY_RUN = True

# Verify-after-submit: after a successful push, poll the listing a few times to
# confirm Amazon accepted it (no ERROR issues) and surface the real outcome into
# the status column instead of just "queued". Adds up to
# VERIFY_MAX_ATTEMPTS × VERIFY_INTERVAL_SECONDS per pushed row.
VERIFY_AFTER_SUBMIT     = True
VERIFY_MAX_ATTEMPTS     = 3
VERIFY_INTERVAL_SECONDS = 10


# =============================================================================
#  CONSTANTS
# =============================================================================

# NOTE: Your existing Amazon listings use product type "SIM_CARD" (confirmed
# from Category Listings Report 05-24-2026). WIRELESS_ACCESSORY is a different
# Amazon product type — change this only if you intentionally want a new type.
PRODUCT_TYPE = "SIM_CARD"

LANGUAGE_TAG    = "en_GB"   # UK English — matches your existing listings
CURRENCY        = "GBP"     # UK marketplace currency
CONDITION_TYPE  = "new_new"
# SIM_CARD deprecated STYLE_NAME; SIZE is a valid single-attribute theme. The
# sheet's "variation" value (e.g. "3GB / 15 Days") is written to the `size` attr.
VARIATION_THEME = "SIZE"

# "DEFAULT" = Merchant Fulfilled, "AMAZON_EU" = FBA (UK/EU).
# eSIMs are digital — they cannot be FBA (Amazon returned warning 12998 when set
# to AMAZON_EU). Merchant Fulfilled is correct: you deliver the QR/activation.
FULFILLMENT_CHANNEL = "DEFAULT"

# ── SIM_CARD required-attribute defaults ────────────────────────────────────
# Amazon's SIM_CARD product type requires these compliance/product attributes.
# Values were chosen to satisfy the schema for a prepaid eSIM (no physical
# battery, no barcode). Adjust if your product details differ.
COUNTRY_OF_ORIGIN   = "GB"              # ISO 3166 alpha-2 (manufacturing origin)
DG_HZ_REGULATION    = "not_applicable"  # Dangerous Goods — eSIM is not regulated
BATTERIES_REQUIRED  = False             # eSIM needs no battery
BATTERIES_INCLUDED  = False
GTIN_EXEMPTION      = True              # eSIMs have no UPC/EAN → claim GTIN exemption
# Nominal package dimensions/weight (eSIM is delivered digitally; Amazon still
# wants package measurements). Centimetres / grams.
PKG_LENGTH_CM       = 12.0
PKG_WIDTH_CM        = 8.0
PKG_HEIGHT_CM       = 1.0
PKG_WEIGHT_GRAMS    = 20.0

# Other SIM_CARD required attributes (validated against the live schema).
NUMBER_OF_ITEMS     = 1
IS_REFURBISHED      = False
POWER_PLUG          = "no_plug"                          # eSIM has no mains plug
GDPR_RISK           = "no_electronic_information_stored"  # SIM stores no personal data
WARRANTY_DESCRIPTION = "No warranty — digital eSIM data service"

# Marketplace ID string derived from the enum — "A1F83G8C2ARO7P" for UK
MARKETPLACE_ID: str = MARKETPLACE.marketplace_id

# ── Google Sheet column indices — 0-based, must match the sheet header order ─
COL_SKU            = 0   # A — child SKU
COL_PARENT_SKU     = 1   # B — parent SKU (shared across variation siblings)
COL_PARENT_TITLE   = 2   # C — title for the non-purchasable parent listing
COL_TITLE          = 3   # D — full child listing title
COL_VARIATION      = 4   # E — variation label, e.g. "3GB / 15 Days"
COL_BRAND          = 5   # F — brand name, e.g. "YouGuide"
COL_PRICE          = 6   # G — selling price (GBP, no £ symbol)
COL_QUANTITY       = 7   # H — stock quantity
COL_DESCRIPTION    = 8   # I — product description
COL_IMAGE1         = 9   # J — main image HTTPS URL
COL_UPLOADED       = 10  # K — written by this script: "YES" on success
COL_STATUS         = 11  # L — written by this script: "UPLOADED"/"UPDATED"/"DELETED"/error
COL_ASIN           = 12  # M — written by this script: Amazon ASIN
COL_HASH           = 13  # N — written by this script: sheet-content hash (all fields)
COL_ACTION         = 14  # O — user-controlled: DELETE/X/TRUE removes; FALSE/blank keeps live
# ── Two-way-sync bookkeeping (written by this script) ───────────────────────
COL_AMAZON_HASH    = 15  # P — robust hash of Amazon-side values at last sync
COL_AMAZON_UPDATED = 16  # Q — Amazon's lastUpdatedDate captured at last sync
COL_SHEET_UPDATED  = 17  # R — when the sheet last pushed a change to Amazon
COL_LAST_SYNCED    = 18  # S — when this row was last reconciled
COL_CONFLICT       = 19  # T — note when sheet & Amazon both changed (sheet won)

# Header labels for columns the script owns — written to row 1 if missing so a
# fresh sheet self-documents. Index → label.
MANAGED_HEADERS = {
    COL_UPLOADED:       "uploaded",
    COL_STATUS:         "status",
    COL_ASIN:           "asin",
    COL_HASH:           "hash",
    COL_ACTION:         "action",
    COL_AMAZON_HASH:    "amazon_hash",
    COL_AMAZON_UPDATED: "amazon_updated",
    COL_SHEET_UPDATED:  "sheet_updated",
    COL_LAST_SYNCED:    "last_synced",
    COL_CONFLICT:       "conflict",
}

# ── Throttle delays ───────────────────────────────────────────────────────────
# SP-API Listings Items burst quota: 5 requests → steady rate: 2 req/s
# Google Sheets write quota: 60 requests/minute per user
SLEEP_AFTER_PARENT = 3.0   # seconds — parent submissions are infrequent
SLEEP_AFTER_CHILD  = 2.0   # seconds — one child API call per row
SLEEP_AFTER_SHEET  = 0.5   # seconds — one batched write per row
SLEEP_AFTER_READ   = 0.6   # seconds — one Amazon read per managed row

# ── Activity-based poll cadence (req: 5–10 min "based on data") ──────────────
# Poll fast (5 min) right after a pass that changed something, then back off
# toward 10 min while nothing is happening; reset to fast on the next change.
POLL_MIN_SECONDS     = 300   # 5 minutes — used after an active pass
POLL_MAX_SECONDS     = 600   # 10 minutes — ceiling while idle
POLL_BACKOFF_SECONDS = 150   # added to the interval after each idle pass

# Sentinel values in the ACTION column (O) that mean "this listing should NOT be
# live on Amazon". Anything else (blank, FALSE, NO, 0) means "should be live".
DELETE_TRIGGERS = {"DELETE", "TRUE", "X", "YES", "1", "REMOVE"}

# Status values that mean the row currently represents a live (uploaded) listing
LIVE_STATUSES = {"UPLOADED", "UPDATED", "PULLED"}


# =============================================================================
#  LOGGING
# =============================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)


def now_iso() -> str:
    """UTC timestamp, second precision — used for the sheet's *_updated columns."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")


# =============================================================================
#  GOOGLE SHEETS HELPERS
# =============================================================================

def get_worksheet() -> gspread.Worksheet:
    """Open the configured worksheet via OAuth user credentials.

    First run opens a browser to authorise; subsequent runs reuse the cached
    token in OAUTH_TOKEN_FILE. Opens by sheet ID so the title can change freely.
    """
    client = gspread.oauth(
        credentials_filename=OAUTH_CREDENTIALS_FILE,
        authorized_user_filename=OAUTH_TOKEN_FILE,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )
    return client.open_by_key(SPREADSHEET_ID).worksheet(WORKSHEET_NAME)


def ensure_headers(worksheet: gspread.Worksheet) -> None:
    """Make sure the script-owned columns exist and are labelled in row 1.

    Widens the sheet and writes the MANAGED_HEADERS labels for any header cell
    that is currently empty, so a sheet that only has the input columns A–J
    gains the bookkeeping columns automatically.
    """
    needed_cols = max(MANAGED_HEADERS) + 1
    if worksheet.col_count < needed_cols:
        worksheet.add_cols(needed_cols - worksheet.col_count)

    header = worksheet.row_values(1)
    cells = []
    for col, label in MANAGED_HEADERS.items():
        current = header[col] if col < len(header) else ""
        if not current.strip():
            cells.append(gspread.Cell(1, col + 1, label))
    if cells and not DRY_RUN:
        worksheet.update_cells(cells)
        log.info("Initialised %d missing header label(s)", len(cells))


def write_cells(worksheet: gspread.Worksheet, sheet_row: int, updates: dict) -> None:
    """Write several cells of one row in a single batched API call.

    `updates` maps a 0-based column index to its new string value. Batching keeps
    us well under the Google Sheets 60-writes/minute quota even with two-way sync.
    """
    if not updates:
        return
    if DRY_RUN:
        preview = {chr(65 + c): v for c, v in updates.items()}
        log.info("[DRY_RUN] would write row %d: %s", sheet_row, preview)
        return
    try:
        cells = [gspread.Cell(sheet_row, col + 1, "" if v is None else str(v))
                 for col, v in updates.items()]
        worksheet.update_cells(cells)
        time.sleep(SLEEP_AFTER_SHEET)
    except Exception as exc:
        log.warning("Sheet update failed for row %d: %s", sheet_row, exc)


def sheet_content_hash(product: dict) -> str:
    """Hash of EVERY user-editable content field — drives Sheet→Amazon updates.

    Any edit to title, price, quantity, description, image, etc. changes this
    hash and so triggers a push to Amazon (requirement: any column edit syncs).
    """
    payload = json.dumps(
        {
            "title":        product["title"],
            "parent_title": product["parent_title"],
            "variation":    product["variation"],
            "brand":        product["brand"],
            "price":        round(float(product["price"]), 2),
            "quantity":     int(product["quantity"]),
            "description":  product["description"],
            "image1":       product["image1"],
        },
        sort_keys=True,
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()[:16]


def amazon_detect_hash(product: dict) -> str:
    """Robust hash for detecting Amazon-side edits.

    Deliberately covers only fields that round-trip cleanly back from Amazon:
    title, brand, variation, price, quantity. Image URLs are re-hosted by Amazon
    on media-amazon.com and description/bullets are normalised, so including them
    would flag a false "Amazon changed" on every poll. When a real change IS
    detected here, the reverse sync still writes back ALL available fields.
    """
    payload = json.dumps(
        {
            "title":     (product.get("title") or "").strip(),
            "brand":     (product.get("brand") or "").strip(),
            "variation": (product.get("variation") or "").strip(),
            "price":     round(float(product.get("price") or 0), 2),
            "quantity":  int(product.get("quantity") or 0),
        },
        sort_keys=True,
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()[:16]


# =============================================================================
#  SP-API CLIENT
# =============================================================================

def get_listings_client() -> ListingsItems:
    return ListingsItems(
        credentials=SP_API_CREDENTIALS,
        marketplace=MARKETPLACE,
    )


# =============================================================================
#  ATTRIBUTE FORMATTERS
# =============================================================================

def text_attr(value: str) -> list:
    """Localised text attribute structure (item_name, brand, style_name, etc.)."""
    return [{"value": value, "language_tag": LANGUAGE_TAG, "marketplace_id": MARKETPLACE_ID}]


def enum_attr(value) -> list:
    """Enumerated / non-localised attribute structure (condition_type, parentage_level, etc.)."""
    return [{"value": value, "marketplace_id": MARKETPLACE_ID}]


def text_attr_multi(values: list) -> list:
    """Localised text attribute with several values (e.g. multiple bullet points)."""
    return [
        {"value": v, "language_tag": LANGUAGE_TAG, "marketplace_id": MARKETPLACE_ID}
        for v in values if v
    ]


def bullet_points(description: str, variation: str) -> list:
    """Derive bullet points for the listing. Edit/extend as marketing requires."""
    bullets = [
        "Prepaid travel data eSIM with instant QR-code activation",
        f"Data plan: {variation}" if variation else "",
        "No physical SIM — activate digitally on eSIM-compatible devices",
        description,
    ]
    return [b for b in bullets if b][:5]   # Amazon allows up to 5 bullet points


def compliance_attrs(brand: str, description: str, variation: str) -> dict:
    """
    SIM_CARD product-level attributes required by Amazon for BOTH parent and
    child (confirmed against the live product-type schema). Values come from the
    configurable constants near the top of this file.
    """
    return {
        "product_description":                text_attr(description),
        "bullet_point":                       text_attr_multi(bullet_points(description, variation)),
        "manufacturer":                       text_attr(brand),
        "country_of_origin":                  enum_attr(COUNTRY_OF_ORIGIN),
        "supplier_declared_dg_hz_regulation": enum_attr(DG_HZ_REGULATION),
        "batteries_required":                 enum_attr(BATTERIES_REQUIRED),
        "batteries_included":                 enum_attr(BATTERIES_INCLUDED),
    }


# =============================================================================
#  PAYLOAD BUILDERS
# =============================================================================

def build_parent_payload(parent_title: str, brand: str, description: str) -> dict:
    """
    Build the PutListingsItemRequest body for a variation parent.

    Parents are non-purchasable — they carry no price, quantity, or condition.
    Using LISTING_PRODUCT_ONLY tells Amazon not to expect offer attributes, but
    the SIM_CARD type still requires the product-level compliance attributes.
    """
    attributes = {
        "item_name":       text_attr(parent_title),
        "brand":           text_attr(brand),
        "parentage_level": enum_attr("parent"),
        "variation_theme": [{"name": VARIATION_THEME, "marketplace_id": MARKETPLACE_ID}],
    }
    attributes.update(compliance_attrs(brand, description, variation=""))
    return {
        "productType":  PRODUCT_TYPE,
        "requirements": "LISTING_PRODUCT_ONLY",
        "attributes":   attributes,
    }


def build_child_payload(
    parent_sku:  str,
    title:       str,
    brand:       str,
    price:       float,
    quantity:    int,
    description: str,
    variation:   str,
    image_url:   str,
) -> dict:
    """Build the PutListingsItemRequest body for a purchasable variation child."""
    attributes: dict = {
        "item_name":       text_attr(title),
        "brand":           text_attr(brand),
        "parentage_level": enum_attr("child"),
        "child_parent_sku_relationship": [
            {
                "child_relationship_type": "variation",
                "parent_sku":              parent_sku,
                "marketplace_id":          MARKETPLACE_ID,
            }
        ],
        # Variation attribute for the SIZE theme (e.g. "3GB / 15 Days").
        "size":           text_attr(variation),
        "condition_type": enum_attr(CONDITION_TYPE),
        "is_refurbished": enum_attr(IS_REFURBISHED),
        "number_of_items": [{"value": NUMBER_OF_ITEMS, "marketplace_id": MARKETPLACE_ID}],
        "power_plug_type": enum_attr(POWER_PLUG),
        "gdpr_risk":      enum_attr(GDPR_RISK),
        "warranty_description": text_attr(WARRANTY_DESCRIPTION),
        "purchasable_offer": [
            {
                "marketplace_id": MARKETPLACE_ID,
                "currency":       CURRENCY,
                "our_price":      [{"schedule": [{"value_with_tax": price}]}],
            }
        ],
        # List price (RRP/MSRP) — required by SIM_CARD; mirror the selling price.
        "list_price": [
            {"currency": CURRENCY, "value_with_tax": price, "marketplace_id": MARKETPLACE_ID}
        ],
        "fulfillment_availability": [
            {
                "fulfillment_channel_code": FULFILLMENT_CHANNEL,
                "quantity":                 quantity,
                "marketplace_id":           MARKETPLACE_ID,
            }
        ],
        # eSIMs have no barcode → declare a GTIN exemption instead of an EAN/UPC.
        "supplier_declared_has_product_identifier_exemption": enum_attr(GTIN_EXEMPTION),
        # Package dimensions + weight (Amazon requires these even for a digital product).
        "item_package_dimensions": [
            {
                "length": {"value": PKG_LENGTH_CM, "unit": "centimeters"},
                "width":  {"value": PKG_WIDTH_CM,  "unit": "centimeters"},
                "height": {"value": PKG_HEIGHT_CM, "unit": "centimeters"},
                "marketplace_id": MARKETPLACE_ID,
            }
        ],
        "item_package_weight": [
            {"value": PKG_WEIGHT_GRAMS, "unit": "grams", "marketplace_id": MARKETPLACE_ID}
        ],
    }
    # Shared product-level compliance attributes (description, bullets,
    # manufacturer, country of origin, dangerous-goods, batteries).
    attributes.update(compliance_attrs(brand, description, variation))

    # Only attach the image attribute when a valid HTTPS URL is provided.
    # Amazon fetches the image from the URL during listing processing.
    # Add other_product_image_locator_1..8 here if additional images are needed.
    if image_url.startswith("https://"):
        attributes["main_product_image_locator"] = [
            {"media_location": image_url, "marketplace_id": MARKETPLACE_ID}
        ]
    elif image_url:
        log.warning("Image URL must use HTTPS — image skipped for this row: %s", image_url)

    return {
        "productType":  PRODUCT_TYPE,
        "requirements": "LISTING",
        "attributes":   attributes,
    }


# =============================================================================
#  SP-API OPERATIONS — WRITES
# =============================================================================

def create_parent_listing(
    client:     ListingsItems,
    parent_sku: str,
    payload:    dict,
) -> bool:
    """Submit a parent listing via SP-API. Returns True when Amazon accepts it."""
    if DRY_RUN:
        log.info("[DRY_RUN] would PUT parent %s", parent_sku)
        return True
    try:
        resp   = client.put_listings_item(
            sellerId=SELLER_ID,
            sku=parent_sku,
            marketplaceIds=[MARKETPLACE_ID],
            body=payload,
        )
        data   = resp.payload or {}
        status = data.get("status", "UNKNOWN")
        issues = data.get("issues", [])

        if status in ("ACCEPTED", "VALID"):
            log.info("[PARENT OK]   %s — %s", parent_sku, status)
            return True

        log.warning(
            "[PARENT FAIL] %s — %s | issues: %s",
            parent_sku, status,
            [i.get("message", str(i)) for i in issues],
        )
        return False

    except SellingApiException as exc:
        log.error("[PARENT ERR]  %s — %s", parent_sku, exc)
        return False


def create_child_listing(
    client:  ListingsItems,
    sku:     str,
    payload: dict,
) -> tuple[bool, str]:
    """Submit a child listing via SP-API. Returns (success, status_message)."""
    if DRY_RUN:
        log.info("[DRY_RUN] would PUT child %s", sku)
        return True, "UPLOADED"
    try:
        resp   = client.put_listings_item(
            sellerId=SELLER_ID,
            sku=sku,
            marketplaceIds=[MARKETPLACE_ID],
            body=payload,
        )
        data   = resp.payload or {}
        status = data.get("status", "UNKNOWN")
        issues = data.get("issues", [])

        if status in ("ACCEPTED", "VALID"):
            log.info("[CHILD OK]    %s — %s", sku, status)
            return True, "UPLOADED"

        # Show up to 15 issues so the full set of validation errors is visible
        issue_summary = "; ".join(i.get("message", str(i)) for i in issues[:15])
        msg = f"FAILED: {status} | {issue_summary}"
        log.warning("[CHILD FAIL]  %s — %s", sku, msg)
        return False, msg

    except SellingApiException as exc:
        msg = f"FAILED: {exc}"
        log.error("[CHILD ERR]   %s — %s", sku, msg)
        return False, msg


def delete_listing(client: ListingsItems, sku: str) -> tuple[bool, str]:
    """Submit a delete request via SP-API. Returns (success, status_message).

    delete_listings_item removes the seller's SKU from the marketplace; the
    underlying ASIN remains, but it no longer appears as one of your offers.
    """
    if DRY_RUN:
        log.info("[DRY_RUN] would DELETE %s", sku)
        return True, "DELETED"
    try:
        resp   = client.delete_listings_item(
            sellerId=SELLER_ID,
            sku=sku,
            marketplaceIds=[MARKETPLACE_ID],
        )
        data   = resp.payload or {}
        status = data.get("status", "UNKNOWN")
        issues = data.get("issues", [])

        if status in ("ACCEPTED", "VALID"):
            log.info("[DELETE OK]   %s — %s", sku, status)
            return True, "DELETED"

        issue_summary = "; ".join(i.get("message", str(i)) for i in issues[:15])
        msg = f"DELETE FAILED: {status} | {issue_summary}"
        log.warning("[DELETE FAIL] %s — %s", sku, msg)
        return False, msg

    except SellingApiException as exc:
        msg = f"DELETE FAILED: {exc}"
        log.error("[DELETE ERR]  %s — %s", sku, msg)
        return False, msg


# =============================================================================
#  SP-API OPERATIONS — READS (reverse sync: Amazon → Sheet)
# =============================================================================

def _attr_text(attrs: dict, key: str) -> str:
    """First localised value of a list-shaped attribute, or ""."""
    v = attrs.get(key)
    if isinstance(v, list) and v and isinstance(v[0], dict):
        return str(v[0].get("value", "")).strip()
    return ""


def _extract_parent_sku(payload: dict) -> str:
    """Pull the parent SKU from a child's relationships block, if present."""
    for rel_group in payload.get("relationships") or []:
        for rel in rel_group.get("relationships") or []:
            parents = rel.get("parentSkus") or []
            if parents:
                return str(parents[0]).strip()
    return ""


def read_amazon_state(client: ListingsItems, sku: str) -> dict:
    """Read a SKU's live state from Amazon for the reverse sync.

    Returns a dict:
        exists       — True if Amazon knows this SKU (False on 404)
        read_ok      — False if the read itself errored (treat state as unknown)
        asin         — Amazon ASIN
        status       — Amazon's listing status (e.g. "BUYABLE")
        last_updated — Amazon's lastUpdatedDate
        product      — fields in the same shape as parse_row(), or None
    """
    result = {"exists": False, "read_ok": True, "asin": "", "status": "",
              "last_updated": "", "product": None}
    try:
        resp = client.get_listings_item(
            sellerId=SELLER_ID,
            sku=sku,
            marketplaceIds=[MARKETPLACE_ID],
            includedData=["summaries", "attributes", "offers",
                          "fulfillmentAvailability", "relationships"],
        )
    except SellingApiException as exc:
        if getattr(exc, "code", None) == 404:
            return result   # genuinely not on Amazon
        log.warning("Amazon read failed for %s: %s", sku, exc)
        result["read_ok"] = False
        return result

    payload    = resp.payload or {}
    summaries  = payload.get("summaries") or []
    attrs      = payload.get("attributes") or {}
    if not summaries and not attrs:
        return result   # nothing to compare against

    result["exists"] = True
    summary = summaries[0] if summaries else {}
    result["asin"] = summary.get("asin", "")
    st = summary.get("status")
    result["status"] = ",".join(st) if isinstance(st, list) else (st or "")
    result["last_updated"] = summary.get("lastUpdatedDate", "")

    # Quantity — prefer the top-level fulfillmentAvailability block.
    quantity = 0
    fa = payload.get("fulfillmentAvailability") or attrs.get("fulfillment_availability") or []
    if fa and isinstance(fa, list):
        try:
            quantity = int(fa[0].get("quantity", 0) or 0)
        except (ValueError, TypeError):
            quantity = 0

    # Price — prefer the offers block, fall back to the purchasable_offer attr.
    price = 0.0
    offers = payload.get("offers") or []
    if offers and isinstance(offers, list):
        try:
            price = float((offers[0].get("price") or {}).get("amount", 0) or 0)
        except (ValueError, TypeError):
            price = 0.0
    if not price:
        try:
            po = attrs.get("purchasable_offer") or []
            price = float(po[0]["our_price"][0]["schedule"][0]["value_with_tax"])
        except (KeyError, IndexError, ValueError, TypeError):
            price = 0.0

    # Image — Amazon re-hosts on media-amazon.com; report it but it won't match
    # the source URL, so it never participates in change DETECTION.
    image = (summary.get("mainImage") or {}).get("link", "") or _attr_text(attrs, "main_product_image_locator")

    result["product"] = {
        "sku":          sku,
        "parent_sku":   _extract_parent_sku(payload),
        "parent_title": "",   # would require a second read of the parent SKU
        "title":        _attr_text(attrs, "item_name") or summary.get("itemName", ""),
        "variation":    _attr_text(attrs, "size"),
        "brand":        _attr_text(attrs, "brand") or _attr_text(attrs, "manufacturer"),
        "price":        price,
        "quantity":     quantity,
        "description":  _attr_text(attrs, "product_description"),
        "image1":       image,
    }
    return result


def fetch_asin(client: ListingsItems, sku: str) -> str:
    """Look up the ASIN Amazon assigned to a SKU. Returns "" when unavailable."""
    state = read_amazon_state(client, sku)
    return state.get("asin", "")


def verify_listing(client: ListingsItems, sku: str) -> tuple[str, str]:
    """Poll a freshly-submitted listing to confirm Amazon accepted it.

    Returns (status, issues_summary). A submission is processed asynchronously,
    so we poll a few times: if any ERROR-severity issue appears the listing was
    rejected; otherwise we report the live status (e.g. "BUYABLE"). issues_summary
    is "" on a clean accept.
    """
    if DRY_RUN:
        return "DRY_RUN", ""
    last_status = ""
    for attempt in range(1, VERIFY_MAX_ATTEMPTS + 1):
        try:
            resp = client.get_listings_item(
                sellerId=SELLER_ID,
                sku=sku,
                marketplaceIds=[MARKETPLACE_ID],
                includedData=["summaries", "issues"],
            )
        except SellingApiException as exc:
            return "", f"verify read failed: {exc}"

        payload   = resp.payload or {}
        summaries = payload.get("summaries") or []
        issues    = payload.get("issues") or []
        errors    = [i for i in issues if i.get("severity") == "ERROR"]

        status = ""
        if summaries:
            st = summaries[0].get("status")
            status = ",".join(st) if isinstance(st, list) else (st or "")
        last_status = status or last_status

        if errors:
            return status or "INVALID", "; ".join(e.get("message", str(e)) for e in errors[:10])
        if status:                       # has a status, no ERROR issues → accepted
            return status, ""
        if attempt < VERIFY_MAX_ATTEMPTS:
            time.sleep(VERIFY_INTERVAL_SECONDS)

    return last_status or "PENDING", ""  # still processing — not an error


# =============================================================================
#  ROW PARSING
# =============================================================================

def parse_row(raw: list, row_index: int) -> Optional[dict]:
    """
    Convert a raw sheet row into a typed product dict.
    row_index is 0-based relative to the data rows (header excluded).
    Returns None and logs a warning when required fields are absent.
    """
    def cell(col: int, default: str = "") -> str:
        return str(raw[col]).strip() if col < len(raw) else default

    try:
        sku        = cell(COL_SKU)
        parent_sku = cell(COL_PARENT_SKU)
        title      = cell(COL_TITLE)

        if not sku or not parent_sku or not title:
            log.warning(
                "Row %d skipped — required field(s) empty: sku=%r  parent_sku=%r  title=%r",
                row_index + 2, sku, parent_sku, title,
            )
            return None

        return {
            "sku":            sku,
            "parent_sku":     parent_sku,
            "parent_title":   cell(COL_PARENT_TITLE),
            "title":          title,
            "variation":      cell(COL_VARIATION),
            "brand":          cell(COL_BRAND),
            "price":          float(cell(COL_PRICE, "0").replace(",", "")),
            "quantity":       int(float(cell(COL_QUANTITY, "0"))),
            "description":    cell(COL_DESCRIPTION),
            "image1":         cell(COL_IMAGE1),
            "status":         cell(COL_STATUS),
            "asin":           cell(COL_ASIN),
            "stored_hash":    cell(COL_HASH),
            "action":         cell(COL_ACTION),
            "amazon_hash":    cell(COL_AMAZON_HASH),
            "amazon_updated": cell(COL_AMAZON_UPDATED),
        }
    except (ValueError, TypeError) as exc:
        log.warning("Row %d parse error — %s", row_index + 2, exc)
        return None


# =============================================================================
#  SHEET ⇄ AMAZON RECONCILE HELPERS
# =============================================================================

def push_to_amazon(
    client:          ListingsItems,
    product:         dict,
    created_parents: set,
) -> tuple[bool, str]:
    """Create/update the parent (once per iteration) and the child on Amazon."""
    parent_sku = product["parent_sku"]
    if parent_sku not in created_parents:
        parent_payload = build_parent_payload(
            parent_title=product["parent_title"] or product["title"],
            brand=product["brand"],
            description=product["description"],
        )
        create_parent_listing(client, parent_sku, parent_payload)
        created_parents.add(parent_sku)
        time.sleep(SLEEP_AFTER_PARENT)

    child_payload = build_child_payload(
        parent_sku=parent_sku,
        title=product["title"],
        brand=product["brand"],
        price=product["price"],
        quantity=product["quantity"],
        description=product["description"],
        variation=product["variation"],
        image_url=product["image1"],
    )
    return create_child_listing(client, product["sku"], child_payload)


def pull_updates(amazon_product: dict) -> dict:
    """Build the column→value map that mirrors Amazon's values into the sheet.

    Empty values are skipped so a partial Amazon read never blanks a good cell.
    Price is skipped when 0 (almost always a read miss, not a genuine free item);
    quantity is always written because 0 legitimately means out of stock.
    """
    p = amazon_product
    candidates = {
        COL_TITLE:       p.get("title", ""),
        COL_VARIATION:   p.get("variation", ""),
        COL_BRAND:       p.get("brand", ""),
        COL_DESCRIPTION: p.get("description", ""),
        COL_IMAGE1:      p.get("image1", ""),
        COL_PARENT_SKU:  p.get("parent_sku", ""),
    }
    updates = {col: val for col, val in candidates.items() if str(val).strip()}
    if float(p.get("price") or 0) > 0:
        updates[COL_PRICE] = p["price"]
    updates[COL_QUANTITY] = int(p.get("quantity") or 0)
    return updates


def record_push_success(
    worksheet:      gspread.Worksheet,
    sheet_row:      int,
    client:         ListingsItems,
    product:        dict,
    sheet_hash_now: str,
    base_status:    str,
    conflict_note:  str = "",
) -> str:
    """Write a successful push back to the sheet, optionally verifying first.

    Returns the final status string written (so the caller can log it).
    """
    sku  = product["sku"]
    asin = product["asin"] or fetch_asin(client, sku)
    status = base_status
    if VERIFY_AFTER_SUBMIT:
        v_status, v_issues = verify_listing(client, sku)
        if v_issues:
            status = f"{base_status} ⚠ {v_issues}"[:400]
        elif v_status:
            status = f"{base_status} ({v_status})"
    write_cells(worksheet, sheet_row, {
        COL_UPLOADED:      "YES",
        COL_STATUS:        status,
        COL_ASIN:          asin or "",
        COL_HASH:          sheet_hash_now,
        COL_AMAZON_HASH:   amazon_detect_hash(product),
        COL_SHEET_UPDATED: now_iso(),
        COL_LAST_SYNCED:   now_iso(),
        COL_CONFLICT:      conflict_note,
    })
    return status


# =============================================================================
#  MAIN ORCHESTRATION
# =============================================================================

def process_sheet(
    worksheet: gspread.Worksheet,
    client:    ListingsItems,
    iteration: int,
) -> int:
    """One full reconcile pass. Returns the number of mutations made this pass
    (used to drive the activity-based poll interval)."""
    all_values = worksheet.get_all_values()

    if len(all_values) < 2:
        log.info("Iteration %d: sheet has no data rows — nothing to do", iteration)
        return 0

    data_rows = all_values[1:]  # strip header row
    log.info("Iteration %d: loaded %d data rows", iteration, len(data_rows))

    created_parents: set = set()        # parent SKUs pushed this iteration
    processed_skus:  dict = {}          # sku → sheet_row of first occurrence
    created = updated = deleted = pulled = conflicts = failed = skipped = 0

    for row_index, raw_row in enumerate(data_rows):
        if not any(c.strip() for c in raw_row):
            continue   # blank row

        sheet_row = row_index + 2   # 1-based sheet row (row 2 = first data row)

        def raw_cell(col: int) -> str:
            return str(raw_row[col]).strip() if col < len(raw_row) else ""

        action_raw     = raw_cell(COL_ACTION).upper()
        current_status = raw_cell(COL_STATUS).upper()
        sku_raw        = raw_cell(COL_SKU)
        want_deleted   = action_raw in DELETE_TRIGGERS

        # ── 1. DELETE desired (declarative): ensure the listing is gone ───────
        if want_deleted:
            if current_status.startswith("DELETED"):
                skipped += 1
                continue   # already deleted; leave action as-is (it's a toggle)
            if not sku_raw:
                log.warning("Row %d: DELETE requested but SKU is empty — ignoring", sheet_row)
                failed += 1
                continue
            log.info("Row %d (%s): DELETE requested", sheet_row, sku_raw)
            success, status_msg = delete_listing(client, sku_raw)
            if success:
                write_cells(worksheet, sheet_row, {
                    COL_UPLOADED:    "",
                    COL_STATUS:      "DELETED",
                    COL_HASH:        "",
                    COL_AMAZON_HASH: "",
                    COL_LAST_SYNCED: now_iso(),
                })
                deleted += 1
            else:
                write_cells(worksheet, sheet_row, {COL_STATUS: status_msg})
                failed += 1
            time.sleep(SLEEP_AFTER_CHILD)
            continue

        product = parse_row(raw_row, row_index)
        if product is None:
            failed += 1
            continue

        # HOLD freezes a row during testing without uploading it.
        if current_status == "HOLD":
            log.info("Row %d (%s): on hold — skipping", sheet_row, product["sku"])
            skipped += 1
            continue

        # Guard against the same SKU appearing more than once in the sheet.
        sku = product["sku"]
        if sku in processed_skus:
            msg = f"SKIPPED: duplicate of row {processed_skus[sku]} (same SKU this iteration)"
            log.warning("Row %d (%s): %s", sheet_row, sku, msg)
            write_cells(worksheet, sheet_row, {COL_STATUS: msg})
            skipped += 1
            continue
        processed_skus[sku] = sheet_row

        # ── Read live Amazon state (needed for both sync directions) ──────────
        amazon = read_amazon_state(client, sku)
        time.sleep(SLEEP_AFTER_READ)
        if not amazon["read_ok"]:
            # Couldn't read Amazon — skip this row this pass rather than guess.
            log.warning("Row %d (%s): Amazon state unknown — skipping this pass", sheet_row, sku)
            skipped += 1
            continue

        sheet_hash_now = sheet_content_hash(product)
        sheet_changed  = sheet_hash_now != product["stored_hash"]
        ever_uploaded  = bool(product["asin"]) or any(
            current_status.startswith(s) for s in LIVE_STATUSES
        )

        # ── 2. Re-add: a DELETED row whose action is now FALSE/blank ──────────
        if current_status.startswith("DELETED"):
            log.info("Row %d (%s): re-create (action cleared on a deleted row)", sheet_row, sku)
            success, status_msg = push_to_amazon(client, product, created_parents)
            if success:
                record_push_success(worksheet, sheet_row, client, product, sheet_hash_now, "UPLOADED")
                created += 1
            else:
                write_cells(worksheet, sheet_row, {COL_STATUS: status_msg})
                failed += 1
            time.sleep(SLEEP_AFTER_CHILD)
            continue

        # ── 3. Amazon-side deletion: was live, now gone on Amazon ─────────────
        if ever_uploaded and not amazon["exists"]:
            log.info("Row %d (%s): listing removed on Amazon — reflecting in sheet", sheet_row, sku)
            write_cells(worksheet, sheet_row, {
                COL_UPLOADED:    "",
                COL_STATUS:      "DELETED (on Amazon)",
                COL_AMAZON_HASH: "",
                COL_LAST_SYNCED: now_iso(),
            })
            deleted += 1
            continue

        # ── 4. Brand-new row: create on Amazon ────────────────────────────────
        if not ever_uploaded:
            log.info("Row %d (%s): CREATE", sheet_row, sku)
            success, status_msg = push_to_amazon(client, product, created_parents)
            if success:
                record_push_success(worksheet, sheet_row, client, product, sheet_hash_now, "UPLOADED")
                created += 1
            else:
                write_cells(worksheet, sheet_row, {COL_STATUS: status_msg, COL_UPLOADED: ""})
                failed += 1
            time.sleep(SLEEP_AFTER_CHILD)
            continue

        # ── 5. Established row: reconcile sheet vs Amazon ─────────────────────
        amazon_hash_now = amazon_detect_hash(amazon["product"]) if amazon["product"] else ""
        amazon_changed  = bool(product["amazon_hash"]) and amazon_hash_now != product["amazon_hash"]

        if sheet_changed and not amazon_changed:
            # Sheet edited → push to Amazon.
            log.info("Row %d (%s): UPDATE (sheet edited)", sheet_row, sku)
            success, status_msg = push_to_amazon(client, product, created_parents)
            if success:
                record_push_success(worksheet, sheet_row, client, product, sheet_hash_now, "UPDATED")
                updated += 1
            else:
                write_cells(worksheet, sheet_row, {COL_STATUS: status_msg})
                failed += 1
            time.sleep(SLEEP_AFTER_CHILD)
            continue

        if amazon_changed and not sheet_changed:
            # Amazon edited externally → pull into the sheet.
            log.info("Row %d (%s): PULL (changed on Amazon)", sheet_row, sku)
            updates = pull_updates(amazon["product"])
            # Recompute the sheet content hash from the merged (post-pull) values
            # so this pull doesn't look like a fresh sheet edit next pass.
            merged = dict(product)
            for col, val in updates.items():
                key = {
                    COL_TITLE: "title", COL_VARIATION: "variation", COL_BRAND: "brand",
                    COL_DESCRIPTION: "description", COL_IMAGE1: "image1",
                    COL_PARENT_SKU: "parent_sku", COL_PRICE: "price", COL_QUANTITY: "quantity",
                }.get(col)
                if key:
                    merged[key] = val
            updates.update({
                COL_STATUS:         "PULLED",
                COL_ASIN:           amazon["asin"] or product["asin"] or "",
                COL_HASH:           sheet_content_hash(merged),
                COL_AMAZON_HASH:    amazon_hash_now,
                COL_AMAZON_UPDATED: amazon["last_updated"],
                COL_LAST_SYNCED:    now_iso(),
                COL_CONFLICT:       "",
            })
            write_cells(worksheet, sheet_row, updates)
            pulled += 1
            continue

        if sheet_changed and amazon_changed:
            # Both moved since last sync → sheet wins; record Amazon's value.
            ap = amazon["product"] or {}
            note = (f"CONFLICT {now_iso()}: sheet kept. Amazon had "
                    f"price={ap.get('price')} qty={ap.get('quantity')} "
                    f"title={ap.get('title')!r} (amazonUpdated {amazon['last_updated']})")
            log.warning("Row %d (%s): %s", sheet_row, sku, note)
            success, status_msg = push_to_amazon(client, product, created_parents)
            if success:
                record_push_success(worksheet, sheet_row, client, product, sheet_hash_now,
                                    "UPDATED", conflict_note=note)
                conflicts += 1
            else:
                write_cells(worksheet, sheet_row, {COL_STATUS: status_msg, COL_CONFLICT: note})
                failed += 1
            time.sleep(SLEEP_AFTER_CHILD)
            continue

        # Nothing changed on either side — keep the row's stored Amazon snapshot
        # current (cheap, single batched write) and move on.
        if not product["amazon_hash"] and amazon_hash_now:
            write_cells(worksheet, sheet_row, {
                COL_AMAZON_HASH:    amazon_hash_now,
                COL_AMAZON_UPDATED: amazon["last_updated"],
                COL_LAST_SYNCED:    now_iso(),
            })
        skipped += 1

    total_changes = created + updated + deleted + pulled + conflicts
    log.info(
        "Iteration %d complete — created:%d updated:%d pulled:%d deleted:%d conflicts:%d failed:%d skipped:%d",
        iteration, created, updated, pulled, deleted, conflicts, failed, skipped,
    )
    return total_changes


def run() -> None:
    log.info("=== Amazon SP-API two-way sync starting%s ===", " [DRY_RUN]" if DRY_RUN else "")
    log.info("Marketplace: %s (%s) | Product type: %s", MARKETPLACE.name, MARKETPLACE_ID, PRODUCT_TYPE)
    log.info("Poll cadence: %d–%ds (activity-based) | Delete triggers: %s",
             POLL_MIN_SECONDS, POLL_MAX_SECONDS, sorted(DELETE_TRIGGERS))
    log.info("Connecting to Google Sheets: id=%s / %r", SPREADSHEET_ID, WORKSHEET_NAME)

    worksheet = get_worksheet()
    client    = get_listings_client()
    ensure_headers(worksheet)

    iteration = 0
    interval  = POLL_MIN_SECONDS
    while True:
        iteration += 1
        log.info("--- Starting iteration %d ---", iteration)
        try:
            changes = process_sheet(worksheet, client, iteration)
        except KeyboardInterrupt:
            log.info("Interrupted by user — exiting")
            return
        except Exception:
            # Catch-all so a transient sheet/API hiccup doesn't kill the loop.
            log.exception("Iteration %d crashed — continuing after sleep", iteration)
            changes = 0

        # Activity-based cadence: fast after a change, back off while idle.
        if changes > 0:
            interval = POLL_MIN_SECONDS
        else:
            interval = min(POLL_MAX_SECONDS, interval + POLL_BACKOFF_SECONDS)

        log.info("Sleeping %ds until next poll…", interval)
        try:
            time.sleep(interval)
        except KeyboardInterrupt:
            log.info("Interrupted by user — exiting")
            return


if __name__ == "__main__":
    run()
