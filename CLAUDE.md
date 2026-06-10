# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A set of standalone Python scripts that sync Amazon listings for YouGuide travel eSIMs / SIM cards between a Google Sheet and Amazon's Selling Partner API (SP-API), targeting the **Amazon UK** marketplace. There is no package, build step, or test suite — each script is run directly.

## Commands

```powershell
pip install python-amazon-sp-api gspread google-auth   # one-time dependency install

python upload_listings.py    # main loop: read sheet → create/update/delete Amazon listings (runs forever, polls every 60s)
python fetch_listings.py     # read-only: pull EVERY live listing via Reports API into the "amazon_current" tab
python fetch_schema.py       # refresh sim_card_schema_*.json from Amazon's live SIM_CARD product-type schema
python diagnose.py           # auth troubleshooter — verifies token + Product Listing role without changing listings
```

There are no lint/test commands. To syntax-check before running:
`python -c "import ast; ast.parse(open('upload_listings.py', encoding='utf-8').read())"`

First run of any script opens a browser for Google OAuth and caches the token to `authorized_user.json`. `upload_listings.py` runs in an infinite polling loop — stop it with Ctrl-C.

## Architecture

`upload_listings.py` is the core; the other three scripts import its config/constants so **credentials and IDs live in exactly one place** (the `SP_API_CREDENTIALS`, `SELLER_ID`, `MARKETPLACE`, `SPREADSHEET_ID` block near the top of `upload_listings.py`). Change config there, not in the dependent scripts.

**Data flow (uploader):** Google Sheet tab `amazon_listings_template` is the source of truth. `process_sheet()` runs one pass per poll:
1. Each data row → `parse_row()` → typed product dict.
2. **Action column (O)** checked first — values in `DELETE_TRIGGERS` (`DELETE`/`TRUE`/`X`/`YES`) trigger `delete_listings_item`.
3. **Change detection** — `row_hash()` (SHA-256 of title/price/qty/description/etc.) is compared with the hash stored in column N. Unchanged + already-`UPLOADED` rows are skipped; changed rows take the UPDATE path.
4. **Parent before child** — one non-purchasable variation parent (`build_parent_payload`, `requirements=LISTING_PRODUCT_ONLY`) is submitted once per unique `parent_sku` per iteration, then each purchasable child (`build_child_payload`, `requirements=LISTING`). `put_listings_item` is an upsert, so the same call creates and updates.
5. Outcome written back to the sheet (columns K–N) via `update_row_status()`.

**Column contract:** `COL_*` constants (0-based) map sheet columns to fields and **must match the sheet's header order**. Columns A–J are user input; K (uploaded), L (status), M (asin), N (hash) are written by the script; O (action) is user-controlled. Sheet rows are 1-based with row 1 = header, so `sheet_row = data_row_index + 2`.

**Status sentinels** in column L drive control flow: `UPLOADED`/`UPDATED` (success), `DELETED` (stays deleted until action cleared), `HOLD` (manually freeze a row — skipped without uploading). Re-running is safe and idempotent thanks to the hash check.

**Attribute formatters** (`text_attr`, `enum_attr`, `text_attr_multi`) wrap values in the localized `{value, language_tag, marketplace_id}` structures SP-API requires. `compliance_attrs()` holds the SIM_CARD product-level attributes Amazon mandates on both parent and child.

**`fetch_listings.py`** deliberately uses the **Reports API** (`GET_MERCHANT_LISTINGS_ALL_DATA`), not `searchListingsItems`, because for this account searchListingsItems returns only one page with no pagination token. It writes to a separate tab (`amazon_current`) and never touches the uploader's tab.

## Domain constraints (already encoded — don't "fix" without reason)

- **Product type is `SIM_CARD`**, not WIRELESS_ACCESSORY — confirmed against the live Category Listings Report. The `sim_card_schema_*.json` files are the authoritative attribute reference; regenerate with `fetch_schema.py`.
- **eSIMs are digital**: fulfillment must be `DEFAULT` (Merchant Fulfilled) — FBA (`AMAZON_EU`) returns warning 12998. GTIN exemption is claimed (no barcode). Variation theme is `SIZE` (SIM_CARD deprecated STYLE_NAME); the sheet's variation label is written to the `size` attribute.
- Marketplace/currency/language are UK-specific: `Marketplaces.UK`, `GBP`, `en_GB`.
- Throttle sleeps (`SLEEP_AFTER_*`) respect SP-API and Google Sheets quotas — don't remove them.

## Security note

`upload_listings.py` currently contains **live SP-API secrets hard-coded** (refresh token, LWA client secret). `oauth_client.json` and `authorized_user.json` hold Google OAuth secrets/tokens. Treat all of these as sensitive; do not commit them to a public remote or echo them in output.
