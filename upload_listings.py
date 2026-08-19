#!/usr/bin/env python3
"""
Amazon SP-API ⇄ Google Sheet two-way sync — YouGuide eSIM / SIM Cards
=====================================================================
Keeps a Google Sheet and your Amazon listings in sync, both directions, on a
continuous activity-based poll. Multi-country: ONE sheet row = one product, and
its countries column (V) names where to sell it — a single code (UK), a comma
list (UK,DE,FR), a region (EU/NA/FE) or ALL (every marketplace: UK/IE/FR/DE/IT/
ES/NL/BE/SE/PL/US/CA/MX/JP/AU). The script pushes that one row's listing (the
SAME SKU) to every country listed; NO extra rows are created. Per-market state
(asin + content hash) is stored compactly in column W so markets already live at
the current content are skipped without an API call, and column L shows a
per-market summary (e.g. "OK:DE,UK | RETRY:JP"). Listing language/currency
follow the marketplace; text is used as-is (no translation) but the PRICE IS
PRICED PER MARKET — the sheet's price column is a BASE price in GBP, marked up
by the competitive tier formula (see "RETAIL PRICING") and then converted into
each marketplace's own currency at today's live rate (see "PRICE CURRENCY
CONVERSION"). One
refresh token per Seller Central authorization group (.env keys
SP_API_REFRESH_TOKEN_EU/NA/JP/AU; JP and AU are issued separately even though
both use the FE endpoint):

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
    python upload_listings.py            # run the two-way sync loop
    python upload_listings.py --rates [BASE_PRICE]
                                         # show the tier markup + today's FX rates:
                                         # what a base price (default 29.99 GBP)
                                         # sells for in every marketplace.
                                         # Read-only — writes nothing.
"""

import hashlib
import json
import logging
import math
import os
import sys
import time
import urllib.error
import urllib.request
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
# Develop Apps → your app → View credentials). The LWA app id + secret are the
# SAME across every region (it is one app); only the refresh token differs per
# region, because the seller authorises the app once per SP-API region.
LWA_CREDENTIALS = {
    "lwa_app_id":        _require("LWA_APP_ID"),
    "lwa_client_secret": _require("LWA_CLIENT_SECRET"),
}

# One refresh token per AUTHORIZATION GROUP, as shown on Seller Central →
# Develop Apps → "Add Authorizations". Each group's "Authorize app" button mints
# a token covering exactly that group's marketplaces:
#   EU → UK, DE, FR, IT, ES, NL, BE, SE, PL, IE   (one token for all of them)
#   NA → US, CA, MX                                (one token)
#   JP → Japan      ─┐ both are the FE *endpoint*, but Seller Central issues a
#   AU → Australia  ─┘ SEPARATE token for each, so they are separate groups here.
# The original single-token var SP_API_REFRESH_TOKEN is still accepted as the EU
# token, and SP_API_REFRESH_TOKEN_FE (if set) is accepted for both JP and AU, so
# existing .env files keep working. Missing tokens → those rows are skipped.
REFRESH_TOKENS = {
    "EU": os.getenv("SP_API_REFRESH_TOKEN_EU") or _require("SP_API_REFRESH_TOKEN"),
    "NA": os.getenv("SP_API_REFRESH_TOKEN_NA", ""),
    "JP": os.getenv("SP_API_REFRESH_TOKEN_JP") or os.getenv("SP_API_REFRESH_TOKEN_FE", ""),
    "AU": os.getenv("SP_API_REFRESH_TOKEN_AU") or os.getenv("SP_API_REFRESH_TOKEN_FE", ""),
}

# Backward-compatible single-credential dict (EU/default token). The helper
# scripts that import this module (diagnose.py, fetch_listings.py, fetch_schema.py,
# inspect_listing.py) use this and MARKETPLACE/MARKETPLACE_ID directly.
SP_API_CREDENTIALS = {
    "refresh_token": REFRESH_TOKENS["EU"],
    **LWA_CREDENTIALS,
}

# Your Amazon Seller account ID  (Seller Central → Settings → Account Info).
# IMPORTANT: the merchant/seller ID is DIFFERENT per SP-API region — the NA and
# FE accounts have their own IDs, distinct from EU. Sending the EU id to a NA/FE
# call returns "Invalid 'sellerId' provided". So, like the tokens, the seller id
# is per authorization group. SELLER_ID (the original var) is the EU id; set
# SELLER_ID_NA / SELLER_ID_JP / SELLER_ID_AU in .env for the other groups.
SELLER_ID = _require("SELLER_ID")
SELLER_IDS = {
    "EU": os.getenv("SELLER_ID_EU") or SELLER_ID,
    "NA": os.getenv("SELLER_ID_NA", ""),
    "JP": os.getenv("SELLER_ID_JP") or os.getenv("SELLER_ID_FE", ""),
    "AU": os.getenv("SELLER_ID_AU") or os.getenv("SELLER_ID_FE", ""),
}

# List every marketplace in ENGLISH (item text uses the per-market English
# locale below instead of the local language). Your content is English, so this
# keeps Amazon from expecting/flagging a local-language listing.
ENGLISH_ONLY = True

# Endpoint marketplace per region. A ListingsItems client only needs *any*
# marketplace from the region to resolve the right endpoint host + signing
# region; the marketplace each call actually targets is passed explicitly as
# marketplaceIds, so one client per region serves all its marketplaces.
REGION_ENDPOINT = {
    "EU": Marketplaces.UK,
    "NA": Marketplaces.US,
    "FE": Marketplaces.JP,
}

# Default marketplace — existing YouGuide listings are on Amazon UK. Kept for the
# helper scripts and used as the fallback when a sheet row's marketplace cell is
# blank (so legacy single-marketplace sheets keep behaving exactly as before).
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
# LIVE: writes are sent to Amazon and the sheet for real.
DRY_RUN = False

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

# NOTE: LANGUAGE_TAG, CURRENCY and MARKETPLACE_ID below are the "active
# marketplace" globals. They default to UK and are repointed per sheet row by
# set_active_marketplace() (processing is single-threaded, so this is safe).
LANGUAGE_TAG    = "en_GB"   # UK English — default; repointed per row
CURRENCY        = "GBP"     # UK marketplace currency — default; repointed per row
CONDITION_TYPE  = "new_new"
# SIM_CARD deprecated STYLE_NAME; SIZE is a valid single-attribute theme. The
# sheet's "variation" value (e.g. "3GB / 15 Days") is written to the `size` attr.
VARIATION_THEME = "SIZE"

# ── Fulfilment mode ──────────────────────────────────────────────────────────
# USE_FBA switches the whole catalogue between Merchant Fulfilled and FBA:
#   False → "DEFAULT"  — merchant fulfilled (FBM); you deliver the QR/activation.
#   True  → the marketplace's Amazon network (AMAZON_EU on UK/EU, etc.) — Amazon
#           ships printed cards from stock held at a fulfilment centre.
#
# History: this was pinned to DEFAULT with a note that eSIMs "cannot be FBA
# (warning 12998)". That no longer holds — YouGuide now prints physical cards and
# stocks them at Amazon, and on 2026-07-17 Amazon accepted AMAZON_EU on the UK
# SIM_CARD listings with no 12998 warning (the UK schema enumerates AMAZON_EU as
# valid). fbm_to_fba.py converts the EXISTING listings; this flag stops the sync
# loop from reverting them, since put_listings_item is a full upsert and would
# otherwise push DEFAULT straight back over the conversion.
USE_FBA = True

# Amazon fulfilment network per SP-API region. NOT keyed by region alone: JP and
# AU share the FE endpoint but are separate fulfilment networks, so AU overrides.
FBA_CHANNEL_BY_REGION = {"EU": "AMAZON_EU", "NA": "AMAZON_NA", "FE": "AMAZON_JP"}
FBA_CHANNEL_BY_MARKET = {"AU": "AMAZON_AU"}
FBM_CHANNEL = "DEFAULT"


def fba_channel_for(code: str) -> str:
    """The Amazon fulfilment network for a marketplace code.

    Only AMAZON_EU is confirmed against a live schema (UK SIM_CARD, 2026-07-17);
    the rest follow Amazon's per-region network naming. A wrong code is rejected
    by the listing schema rather than applied silently, so it surfaces loudly.
    """
    cfg = MARKETPLACES[code]
    return FBA_CHANNEL_BY_MARKET.get(code, FBA_CHANNEL_BY_REGION[cfg["region"]])

# Active value — repointed per row by set_active_marketplace(), because the right
# FBA network depends on the marketplace being written.
FULFILLMENT_CHANNEL = "DEFAULT"

# ── SIM_CARD required-attribute defaults ────────────────────────────────────
# Amazon's SIM_CARD product type requires these compliance/product attributes.
# Values were chosen to satisfy the schema for a prepaid eSIM (no physical
# battery, no barcode). Adjust if your product details differ.
COUNTRY_OF_ORIGIN   = "GB"              # ISO 3166 alpha-2 (manufacturing origin)
DG_HZ_REGULATION    = "not_applicable"  # Dangerous Goods — eSIM is not regulated
BATTERIES_REQUIRED  = False             # eSIM needs no battery
BATTERIES_INCLUDED  = False
GTIN_EXEMPTION      = True              # fallback when a row has no product_id (col U)
PRODUCT_ID_TYPE     = "ean"             # col U barcodes are EANs (enum: ean/gtin/upc)
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

# Marketplace ID string derived from the enum — "A1F83G8C2ARO7P" for UK.
# This is the DEFAULT/active value; set_active_marketplace() repoints it per row.
MARKETPLACE_ID: str = MARKETPLACE.marketplace_id

# ── Marketplace registry (multi-marketplace support) ─────────────────────────
# Each sheet row names a marketplace in column V (see COL_MARKETPLACE). The code
# here maps that short code → Amazon marketplace id, listing language, currency
# and SP-API region (which selects the refresh token + endpoint). Add/remove
# rows here to control which marketplaces the sync targets. Tokens you have not
# configured simply cause their region's marketplaces to be skipped.
#
# Listing language follows YouGuide's per-country requirement (Belgium prints
# Dutch front / French back → listed in Dutch). Manufacturing origin
# (COUNTRY_OF_ORIGIN) is identical everywhere, so it is not part of this table.
# "region" selects the SP-API endpoint host; "token" selects the refresh-token
# authorization group (see REFRESH_TOKENS). They match except for FE, where JP
# and AU share the FE endpoint but each has its own token.
# "lang" = native locale. "en" = the locale tag the script actually sends when
# ENGLISH_ONLY is True. Some stores ACCEPT English listings (UK/IE/DE/FR/IT/ES/
# US/CA/AU → English locale), but BE/NL/SE/PL/MX/JP do NOT allow an English item
# name — Amazon requires the LOCAL language there — so those keep their native
# locale (your English text is still sent, just tagged with the local locale).
MARKETPLACES = {
    # code:  marketplace_id,        native,    list-locale, currency, region, token
    "UK": {"id": "A1F83G8C2ARO7P", "lang": "en_GB", "en": "en_GB", "currency": "GBP", "region": "EU", "token": "EU"},
    "IE": {"id": "A28R8C7NBKEWEA", "lang": "en_IE", "en": "en_IE", "currency": "EUR", "region": "EU", "token": "EU"},
    "FR": {"id": "A13V1IB3VIYZZH", "lang": "fr_FR", "en": "en_GB", "currency": "EUR", "region": "EU", "token": "EU"},
    "DE": {"id": "A1PA6795UKMFR9", "lang": "de_DE", "en": "en_GB", "currency": "EUR", "region": "EU", "token": "EU"},
    "IT": {"id": "APJ6JRA9NG5V4",  "lang": "it_IT", "en": "en_GB", "currency": "EUR", "region": "EU", "token": "EU"},
    "ES": {"id": "A1RKKUPIHCS9HS", "lang": "es_ES", "en": "en_GB", "currency": "EUR", "region": "EU", "token": "EU"},
    "NL": {"id": "A1805IZSGTT6HS", "lang": "nl_NL", "en": "nl_NL", "currency": "EUR", "region": "EU", "token": "EU"},
    "BE": {"id": "AMEN7PMS3EDWL",  "lang": "fr_BE", "en": "fr_BE", "currency": "EUR", "region": "EU", "token": "EU"},
    "SE": {"id": "A2NODRKZP88ZB9", "lang": "sv_SE", "en": "sv_SE", "currency": "SEK", "region": "EU", "token": "EU"},
    "PL": {"id": "A1C3SOZRARQ6R3", "lang": "pl_PL", "en": "pl_PL", "currency": "PLN", "region": "EU", "token": "EU"},
    "US": {"id": "ATVPDKIKX0DER",  "lang": "en_US", "en": "en_US", "currency": "USD", "region": "NA", "token": "NA"},
    "CA": {"id": "A2EUQ1WTGCTBG2", "lang": "en_CA", "en": "en_CA", "currency": "CAD", "region": "NA", "token": "NA"},
    "MX": {"id": "A1AM78C64UM0Y8", "lang": "es_MX", "en": "es_MX", "currency": "MXN", "region": "NA", "token": "NA"},
    "JP": {"id": "A1VC38T7YXB528", "lang": "ja_JP", "en": "ja_JP", "currency": "JPY", "region": "FE", "token": "JP"},
    "AU": {"id": "A39IBJ37TRP1C6", "lang": "en_AU", "en": "en_AU", "currency": "AUD", "region": "FE", "token": "AU"},
}

# Marketplace code used for sheet rows that leave the marketplace cell blank —
# keeps legacy UK-only sheets working untouched.
DEFAULT_MARKETPLACE_CODE = "UK"

# Resolve the active fulfilment channel for the default marketplace now that the
# registry exists. set_active_marketplace() repoints this per row; this initial
# value keeps any payload built before that call consistent with USE_FBA.
FULFILLMENT_CHANNEL = (fba_channel_for(DEFAULT_MARKETPLACE_CODE) if USE_FBA
                       else FBM_CHANNEL)

# The marketplace the ORIGINAL (master) rows already represent. Every existing
# listing in the sheet is live on UK, so a master row is always processed as its
# HOME (UK) listing on its current SKU — that is how already-uploaded listings
# are preserved (updated in place, never duplicated). The countries listed in
# column V *in addition* to HOME are the ones that get auto-generated rows.
HOME_MARKETPLACE = "UK"

# Country-group shorthands accepted in the countries column (V). "ALL" = every
# marketplace; a region code expands to that region's marketplaces.
COUNTRY_GROUPS = {
    "ALL": list(MARKETPLACES),
    "EU":  [c for c, cfg in MARKETPLACES.items() if cfg["region"] == "EU"],
    "NA":  [c for c, cfg in MARKETPLACES.items() if cfg["region"] == "NA"],
    "FE":  [c for c, cfg in MARKETPLACES.items() if cfg["region"] == "FE"],
}


def parse_countries(text: str) -> tuple[list, list]:
    """Parse a countries-column value into (valid_codes, invalid_tokens).

    Accepts a single code (UK), a comma/space/semicolon list (UK, DE, FR), a
    region/group shorthand (ALL/EU/NA/FE), or any mix. Order is preserved and
    duplicates removed. Unknown tokens are returned separately so the caller can
    surface them instead of silently dropping rows.
    """
    codes: list = []
    invalid: list = []
    for tok in str(text or "").replace(";", ",").replace(" ", ",").split(","):
        tok = tok.strip().upper()
        if not tok:
            continue
        if tok in COUNTRY_GROUPS:
            expansion = COUNTRY_GROUPS[tok]
        elif tok in MARKETPLACES:
            expansion = [tok]
        else:
            invalid.append(tok)
            continue
        for c in expansion:
            if c not in codes:
                codes.append(c)
    return codes, invalid


def set_active_marketplace(code: str) -> None:
    """Repoint the module-level attribute formatters at the given marketplace.

    The payload builders and formatters (text_attr/enum_attr/…) read
    MARKETPLACE_ID, LANGUAGE_TAG and CURRENCY as module globals. Processing is
    single-threaded and sequential, so we simply repoint them before handling
    each row rather than threading a context object through every function.
    """
    global MARKETPLACE_ID, LANGUAGE_TAG, CURRENCY, SELLER_ID, FULFILLMENT_CHANNEL
    cfg = MARKETPLACES[code]
    MARKETPLACE_ID = cfg["id"]
    LANGUAGE_TAG   = cfg["en"] if ENGLISH_ONLY else cfg["lang"]
    CURRENCY       = cfg["currency"]
    # Seller/merchant id is per authorization group (EU/NA/JP/AU).
    SELLER_ID      = SELLER_IDS.get(cfg["token"], "")
    # FBA network is per marketplace — AMAZON_EU is rejected on a US listing.
    FULFILLMENT_CHANNEL = fba_channel_for(code) if USE_FBA else FBM_CHANNEL


# =============================================================================
#  PRICE CURRENCY CONVERSION (live FX)
# =============================================================================
# The sheet's price column (G) holds ONE number per product: the BASE price in
# BASE_CURRENCY (GBP). Two steps turn it into what a shopper sees:
#   1. retail_price_raw() — the tier formula (see RETAIL PRICING below) marks the
#                        cost up to the GBP selling price, left UNROUNDED.
#   2. market_price()  — that selling price is converted into the marketplace's
#                        own currency at today's live rate, and ONLY THEN rounded
#                        up to the next .99. Rounding in GBP first and again after
#                        conversion would stack two round-ups on top of each other
#                        (£6.24 -> £6.99 -> €8.15 -> €8.99 instead of €7.99).
# So a 50.99 base price is listed as 67.99 GBP in the UK and ~79.99 EUR in
# Germany — never as 50.99 PLN in Poland.
#
# Rates come from free, key-less public feeds (tried in order) and are cached in
# memory + on disk for FX_TTL_HOURS, so a poll every 5 minutes does not hammer
# them and a restart doesn't need a fresh fetch. If every feed is unreachable the
# last cached rates are reused (however old) and, failing even that, the affected
# markets are reported NO_FX and SKIPPED — a stale-but-real price is fine, a
# 29.99 PLN price is not.
#
# The converted price is what the per-market hash (column W) is built from, so a
# market re-pushes only when ITS OWN local price actually moves. With charm
# rounding a rate has to drift ~1-3% before the rounded price changes at all,
# which keeps the sync quiet instead of re-pushing the catalogue every day.
#
# CHANGING BASE_CURRENCY REPRICES LISTINGS. Every market except the new base one
# gets a different number, so the next pass re-pushes them all — check
# `--rates <a real sheet price>` before restarting the loop.

CONVERT_PRICE  = True          # False → send the sheet number as-is everywhere
# Currency of the sheet's price column (G). Column G holds the WHOLESALE COST in
# USD, as supplied by the eSIM Access API (fetch_esim_prices.py writes it) — not
# the shelf price. The tier formula turns it into a selling price and only then
# is it converted per marketplace.
BASE_CURRENCY  = (os.getenv("BASE_CURRENCY", "USD") or "USD").strip().upper()

# Currency the pricing policy is expressed in: the tier thresholds (£5/£15/£40)
# and MIN_PROFIT (£4) are pounds, so the USD cost is converted to GBP, priced,
# and the resulting GBP selling price is then converted out to each marketplace.
PRICING_CURRENCY = (os.getenv("PRICING_CURRENCY", "GBP") or "GBP").strip().upper()

# Uplift applied on top of the pure FX price, in percent. Covers the extra cost
# of selling abroad (referral fees, VAT differences, FX spread). 0 = pure rate.
FX_MARKUP_PCT  = float(os.getenv("FX_MARKUP_PCT", "0") or 0)

# Price rounding after conversion:
#   "charm"   → the next .99 up (34.19 → 34.99); never rounds DOWN, so a
#               converted price is never below the true converted value.
#   "nearest" → plain rounding to the currency's precision (34.19 → 34.19).
FX_ROUNDING    = (os.getenv("FX_ROUNDING", "charm") or "charm").strip().lower()

# Currencies Amazon quotes with no decimal part. JPY prices are whole yen, so
# charm rounding uses a yen-style ending instead (…,980) on a 100-yen grid.
ZERO_DECIMAL_CURRENCIES = {"JPY"}
ZERO_DECIMAL_GRID       = 100
ZERO_DECIMAL_ENDING     = 80

# How long a fetched rate set is considered current, and where it is cached.
FX_TTL_HOURS  = float(os.getenv("FX_TTL_HOURS", "12") or 12)
FX_CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fx_rates.json")
FX_TIMEOUT    = 15   # seconds per feed request

# Manual rate pins — set a currency here (or FX_RATE_<CUR> in .env) to bypass the
# live feed for it, e.g. to hold a promotional price point steady.
FX_RATE_OVERRIDES: dict = {
    cur: float(os.environ["FX_RATE_" + cur])
    for cur in {cfg["currency"] for cfg in MARKETPLACES.values()}
    if os.getenv("FX_RATE_" + cur)
}


def _fx_parse_er_api(data: dict) -> Optional[dict]:
    """open.er-api.com — {"result":"success","rates":{"EUR":1.17,…}}."""
    if data.get("result") != "success":
        return None
    return data.get("rates")


def _fx_parse_frankfurter(data: dict) -> Optional[dict]:
    """frankfurter.app (ECB reference rates) — {"rates":{"EUR":1.17,…}}."""
    return data.get("rates")


def _fx_parse_currency_api(data: dict) -> Optional[dict]:
    """fawazahmed0 currency-api — {"gbp":{"eur":1.17,…}} (lowercase keys)."""
    inner = data.get(BASE_CURRENCY.lower())
    if not isinstance(inner, dict):
        return None
    return {k.upper(): v for k, v in inner.items()}


# (name, url, parser). Tried in order; the first response that covers every
# currency the registry needs wins. All three are free and need no API key.
FX_SOURCES = [
    ("open.er-api.com", "https://open.er-api.com/v6/latest/{base}", _fx_parse_er_api),
    ("frankfurter.app", "https://api.frankfurter.app/latest?from={base}", _fx_parse_frankfurter),
    ("currency-api",
     "https://cdn.jsdelivr.net/npm/@fawazahmed0/currency-api@latest/v1/currencies/{base_lower}.json",
     _fx_parse_currency_api),
]

# In-memory cache: {"rates": {...}, "fetched": epoch_seconds, "source": name}.
_FX_CACHE: dict = {"rates": {}, "fetched": 0.0, "source": ""}


def _needed_currencies() -> set:
    """Every currency the pricing chain can ask for, minus the base.

    The marketplaces' own currencies plus PRICING_CURRENCY — the tier formula runs
    in pounds, so a GBP rate is required even though no rate table lists it twice.
    """
    return ({cfg["currency"] for cfg in MARKETPLACES.values()}
            | {PRICING_CURRENCY}) - {BASE_CURRENCY}


def _fx_http_json(url: str) -> dict:
    req = urllib.request.Request(url, headers={"User-Agent": "youguide-amazon-sync/1.0"})
    with urllib.request.urlopen(req, timeout=FX_TIMEOUT) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _fx_read_disk_cache() -> dict:
    """Last rates written by a previous run — survives restarts and outages."""
    try:
        with open(FX_CACHE_FILE, encoding="utf-8") as fh:
            data = json.load(fh)
        if (data.get("base") == BASE_CURRENCY and isinstance(data.get("rates"), dict)
                and data["rates"]):
            return data
    except (OSError, ValueError, TypeError):
        pass
    return {}


def _fx_write_disk_cache(rates: dict, source: str) -> None:
    try:
        with open(FX_CACHE_FILE, "w", encoding="utf-8") as fh:
            json.dump({"base": BASE_CURRENCY, "rates": rates, "source": source,
                       "fetched": time.time(), "fetched_iso": now_iso()}, fh, indent=2)
    except OSError as exc:
        log.warning("Could not write FX cache %s: %s", FX_CACHE_FILE, exc)


def fx_rates(force: bool = False) -> dict:
    """Rates as {currency: units per 1 BASE_CURRENCY}. {} when totally unavailable.

    Cheap to call repeatedly: it only hits the network when the in-memory copy is
    older than FX_TTL_HOURS (or force=True). Order of preference:
      1. fresh in-memory rates      2. fresh on-disk rates (a restart reuses them)
      3. a live fetch               4. STALE on-disk rates, loudly logged
    Manual pins (FX_RATE_OVERRIDES) are layered on top of whatever wins.
    """
    ttl = FX_TTL_HOURS * 3600
    if not force and _FX_CACHE["rates"] and (time.time() - _FX_CACHE["fetched"]) < ttl:
        return dict(_FX_CACHE["rates"], **FX_RATE_OVERRIDES)

    if not force and not _FX_CACHE["rates"]:
        disk = _fx_read_disk_cache()
        if disk and (time.time() - float(disk.get("fetched") or 0)) < ttl:
            _FX_CACHE.update(rates=disk["rates"], fetched=float(disk["fetched"]),
                             source=disk.get("source", "disk cache"))
            log.info("FX: reusing cached rates from %s (%s)",
                     disk.get("source", "disk"), disk.get("fetched_iso", "unknown time"))
            return dict(_FX_CACHE["rates"], **FX_RATE_OVERRIDES)

    needed = _needed_currencies()
    best: tuple = ({}, "", -1)   # (rates, source, coverage)
    for name, url, parser in FX_SOURCES:
        try:
            raw = _fx_http_json(url.format(base=BASE_CURRENCY, base_lower=BASE_CURRENCY.lower()))
            rates = parser(raw) or {}
            rates = {k: float(v) for k, v in rates.items()
                     if isinstance(v, (int, float)) and float(v) > 0}
        except (urllib.error.URLError, OSError, ValueError, TypeError, KeyError) as exc:
            log.warning("FX: source %s failed (%s) — trying next", name, exc)
            continue
        covered = len(needed & set(rates))
        if covered > best[2]:
            best = (rates, name, covered)
        if covered == len(needed):
            break

    rates, source, covered = best
    if rates:
        if covered < len(needed):
            log.warning("FX: %s is missing rates for %s — those markets will be skipped",
                        source, ", ".join(sorted(needed - set(rates))) or "none")
        _FX_CACHE.update(rates=rates, fetched=time.time(), source=source)
        _fx_write_disk_cache(rates, source)
        log.info("FX: rates refreshed from %s (base %s)", source, BASE_CURRENCY)
        return dict(rates, **FX_RATE_OVERRIDES)

    # Every feed is down — fall back to the last known rates, however old, rather
    # than to no conversion at all (which would list EUR numbers as PLN/JPY).
    disk = _fx_read_disk_cache()
    if disk:
        age_h = (time.time() - float(disk.get("fetched") or 0)) / 3600
        log.warning("FX: all sources unreachable — using cached rates from %s (%.1f h old)",
                    disk.get("fetched_iso", "unknown time"), age_h)
        _FX_CACHE.update(rates=disk["rates"], fetched=float(disk.get("fetched") or 0),
                         source=disk.get("source", "disk cache") + " (stale)")
        return dict(disk["rates"], **FX_RATE_OVERRIDES)

    if FX_RATE_OVERRIDES:
        log.warning("FX: no live or cached rates — only the manually pinned ones are usable")
        return dict(FX_RATE_OVERRIDES)

    log.error("FX: no exchange rates available — non-%s markets will be skipped this pass",
              BASE_CURRENCY)
    return {}


def convert_amount(amount: float, frm: str, to: str) -> Optional[float]:
    """Convert between ANY two currencies, unrounded. None if a rate is missing.

    Rates are quoted per 1 BASE_CURRENCY, so a cross rate goes via the base:
    amount ÷ rate[from] × rate[to]. Used for USD cost → GBP pricing → local shelf
    price, which is two hops that never touch a rounding step in between.
    """
    if frm == to:
        return float(amount)
    rates  = fx_rates()
    r_from = 1.0 if frm == BASE_CURRENCY else rates.get(frm)
    r_to   = 1.0 if to  == BASE_CURRENCY else rates.get(to)
    if not r_from or not r_to or r_from <= 0 or r_to <= 0:
        return None
    return float(amount) / float(r_from) * float(r_to)


def round_price(value: float, currency: str, mode: Optional[str] = None) -> float:
    """Round a converted price to a sellable number in the target currency.

    "charm" never rounds down: the result is the next .99 at or above `value`
    (…,980 on the 100-unit grid for zero-decimal currencies like JPY), so an
    exchange-rate move can never quietly undercut the intended margin.
    """
    mode     = (mode or FX_ROUNDING)
    zero_dec = currency in ZERO_DECIMAL_CURRENCIES
    if mode == "charm":
        if zero_dec:
            steps = math.ceil((value - ZERO_DECIMAL_ENDING) / ZERO_DECIMAL_GRID)
            return float(max(1, steps) * ZERO_DECIMAL_GRID + ZERO_DECIMAL_ENDING)
        candidate = math.floor(value) + 0.99
        if candidate < value - 1e-9:
            candidate += 1.0
        return round(candidate, 2)
    # "nearest" — plain rounding to the currency's precision.
    return float(round(value)) if zero_dec else round(value + 1e-9, 2)


# =============================================================================
#  RETAIL PRICING (competitive tier margins on the base price)
# =============================================================================
# Column G is the WHOLESALE COST in USD (from the eSIM Access API), not the shelf
# price. The selling price is derived from it in PRICING_CURRENCY (GBP) — the cost
# is converted USD→GBP first, because the policy below is written in pounds:
#
#   1. Tier multiplier — cheaper items carry a bigger multiple, because a flat
#      percentage leaves no margin at the bottom of the range:
#          cost <  £5   → ×3.0        £15 ≤ cost < £40 → ×1.6
#          £5 ≤ cost < £15 → ×2.0     cost ≥ £40       → ×1.3
#   2. Minimum-profit safeguard — Amazon takes ~AMAZON_FEE_PCT of the sale, so a
#      thin multiple can still leave less than MIN_PROFIT per unit. The price is
#      lifted to whatever clears £4 profit after fees:
#          profit = price × (1 − fee) − cost  ≥  MIN_PROFIT
#          ⇒ price ≥ (cost + MIN_PROFIT) / (1 − fee)
#      The higher of the two wins, so the safeguard only ever raises a price.
#
# That result is the RAW selling price and is NOT rounded here. The .99 is applied
# once, per marketplace, after the FX conversion (market_price) — so every store
# shows a genuine x.99 in its own currency instead of the converted image of a
# British one. Worked example ($11.25 cost):
#   → £8.55 | tier ×2 → 17.10 | safeguard (8.55+4)/0.82 → 15.30 | max → 17.10 raw
#   → UK £17.99  |  DE 17.10 x1.166 = €19.94 → €19.99  |  PL 85.77 zł → 85.99 zł
#
# The fee rate is a single blended approximation of Amazon's referral fee — it
# does NOT model FBA fulfilment fees, storage, or VAT, so MIN_PROFIT is a floor
# on gross contribution, not true net profit. Raise MIN_PROFIT if you want the
# safeguard to cover fulfilment too.

TIER_PRICING   = True   # False → the cost IS the selling price (FX only)

# Set by --first <text> on the command line: rows whose SKU or title contain this
# are processed before all the others in every pass. "" = plain sheet order.
PRIORITY_FIRST = ""

# (upper bound EXCLUSIVE in PRICING_CURRENCY, multiplier). None = no upper bound.
# Boundaries land in the cheaper tier: exactly £15 is ×1.6, not ×2.0.
PRICE_TIERS = [
    (5.0,  3.0),
    (15.0, 2.0),
    (40.0, 1.6),
    (None, 1.3),
]

AMAZON_FEE_PCT = float(os.getenv("AMAZON_FEE_PCT", "18") or 18)   # blended referral fee, %
MIN_PROFIT     = float(os.getenv("MIN_PROFIT", "4") or 4)         # per sale, in PRICING_CURRENCY


def tier_multiplier(cost: float) -> float:
    """The competitive-tier multiplier for a cost in PRICING_CURRENCY."""
    for upper, mult in PRICE_TIERS:
        if upper is None or cost < upper:
            return mult
    return PRICE_TIERS[-1][1]


def retail_price_raw(cost_in_base: float) -> Optional[float]:
    """USD cost (column G) → UNROUNDED selling price in PRICING_CURRENCY (GBP).

    Tier multiple + minimum-profit safeguard only — deliberately NOT rounded to
    .99, because the psychological price point belongs in the currency the
    shopper actually pays in (see market_price). Rounding here as well would
    compound: £6.24 -> £6.99 -> x1.166 -> €8.15 -> €8.99, where rounding once
    after conversion gives €7.99 for the same cost.

    Returns 0.0 for a non-positive cost, so an empty or garbage cell cannot
    invent a price out of the minimum-profit rule alone, and None when the
    USD→GBP rate is unavailable (the caller then skips the row/market rather than
    guessing).
    """
    cost_base = round(float(cost_in_base or 0), 2)
    if cost_base <= 0:
        return 0.0
    cost = convert_amount(cost_base, BASE_CURRENCY, PRICING_CURRENCY)
    if cost is None:
        return None
    if not TIER_PRICING:
        return round(cost, 2)
    tiered    = cost * tier_multiplier(cost)
    fee_share = 1.0 - (AMAZON_FEE_PCT / 100.0)
    floor     = ((cost + MIN_PROFIT) / fee_share) if fee_share > 0 else tiered
    return max(tiered, floor)


def retail_price(cost_in_base: float) -> Optional[float]:
    """The PRICING_CURRENCY (GBP) shelf price: retail_price_raw() at the next .99.

    This is the UK's own listed price and the headline number in previews and
    logs. Every OTHER marketplace converts the RAW price and rounds in its own
    currency instead — market_price() — so each store gets a true .99 rather
    than the FX image of a British one.
    """
    sell = retail_price_raw(cost_in_base)
    if sell is None or sell <= 0 or not TIER_PRICING:
        return sell   # tier pricing off → the cost IS the price, left as-is
    # Always .99 here regardless of FX_ROUNDING — the psychological price point
    # is part of the pricing policy, not of the currency conversion.
    return round_price(sell, PRICING_CURRENCY, mode="charm")


def price_breakdown(cost_in_base: float) -> dict:
    """Explain one cost end to end — used by --rates and to eyeball the policy."""
    cost_base = round(float(cost_in_base or 0), 2)
    cost      = convert_amount(cost_base, BASE_CURRENCY, PRICING_CURRENCY)
    raw       = retail_price_raw(cost_base)
    sell      = retail_price(cost_base)
    if cost is None or sell is None or raw is None:
        return {"cost_base": cost_base, "cost": None, "multiplier": 0.0,
                "tiered": 0.0, "min_profit_price": 0.0, "sell_raw": None,
                "sell": None, "profit": 0.0, "safeguard_applied": False}
    cost   = round(cost, 2)
    mult   = tier_multiplier(cost)
    fee    = 1.0 - (AMAZON_FEE_PCT / 100.0)
    floor  = (cost + MIN_PROFIT) / fee if fee > 0 else 0.0
    return {
        "cost_base": cost_base, "cost": cost, "multiplier": mult,
        "tiered": round(cost * mult, 2), "min_profit_price": round(floor, 2),
        # sell_raw is what every non-GBP market converts; sell is the UK price.
        "sell_raw": round(raw, 4), "sell": sell,
        "profit": round(sell * fee - cost, 2),
        "safeguard_applied": floor > cost * mult,
    }


def market_price(base_price: float, code: str) -> Optional[float]:
    """USD cost (column G) → the number to list in `code`. None if unavailable.

    The single entry point for "what does this SKU cost a shopper here":
        cost USD → GBP → tier + min-profit → FX → round UP to the local .99
    Note the ORDER: the .99 is applied once, at the end, in the marketplace's own
    currency — so Poland gets a real 31,99 zł rather than the converted image of
    a British £6.99 (which lands on 35,99 zł). Every caller (payload builder,
    per-market hash, plan preview, --rates) goes through this function, so they
    cannot disagree about the price.

    None means a required rate is missing — the caller must SKIP the market
    rather than send a number under the wrong currency code.
    """
    cfg      = MARKETPLACES[code]
    currency = cfg["currency"]
    sell     = retail_price_raw(base_price)  # tier + min-profit, GBP, unrounded
    if sell is None:
        return None
    if sell <= 0:
        return 0.0                           # no cost in column G — caller skips
    if not CONVERT_PRICE or currency == PRICING_CURRENCY:
        return retail_price(base_price)      # the .99 lands in GBP, the pay currency
    local = convert_amount(sell, PRICING_CURRENCY, currency)
    if local is None:
        return None
    return round_price(local * (1 + FX_MARKUP_PCT / 100.0), currency)


def base_price_from(local_price: float, code: str) -> Optional[float]:
    """A marketplace's local price expressed in BASE_CURRENCY (FX inverse only).

    This undoes the currency conversion, NOT the retail tier formula — see
    pull_updates(), which refuses to write the price back at all while
    TIER_PRICING is on, because column G is a cost input the pricing policy reads
    rather than a shelf price to be mirrored. None when the rate is unknown.
    """
    cfg      = MARKETPLACES.get(code) or MARKETPLACES[HOME_MARKETPLACE]
    currency = cfg["currency"]
    if not CONVERT_PRICE or currency == BASE_CURRENCY:
        return round(float(local_price or 0), 2)
    rate = fx_rates().get(currency)
    if not rate or rate <= 0:
        return None
    return round(float(local_price or 0) / float(rate) / (1 + FX_MARKUP_PCT / 100.0), 2)


def log_fx_table(cost: float = 11.25) -> None:
    """Log the whole chain for one USD cost: tiers → GBP price → every market."""
    rates = fx_rates()
    b     = price_breakdown(cost)
    if b["cost"] is None:
        log.error("  Pricing: no %s→%s rate — cannot price anything this pass",
                  BASE_CURRENCY, PRICING_CURRENCY)
        return
    log.info("  Cost (col G)  : %.2f %s  =  %.2f %s", b["cost_base"], BASE_CURRENCY,
             b["cost"], PRICING_CURRENCY)
    log.info("  Tier pricing  : %s", "ON" if TIER_PRICING else "OFF (cost = selling price)")
    if TIER_PRICING:
        log.info("     x%.1f = %.2f  |  min-%.2f-profit floor %.2f%s",
                 b["multiplier"], b["tiered"], MIN_PROFIT, b["min_profit_price"],
                 "  <- applied" if b["safeguard_applied"] else "")
    log.info("     SELLING PRICE %.2f %s -> %.2f %s   (profit %.2f after %.0f%% fees)",
             b["sell_raw"], PRICING_CURRENCY, b["sell"], PRICING_CURRENCY,
             b["profit"], AMAZON_FEE_PCT)
    log.info("     (the UNROUNDED %.2f %s is what other markets convert; each one "
             "rounds up to its own .99)", b["sell_raw"], PRICING_CURRENCY)
    log.info("  FX conversion : %s", "ON" if CONVERT_PRICE else "OFF (selling price sent as-is)")
    if not CONVERT_PRICE:
        return
    log.info("  FX base/round : rates per 1 %s / %s%s", BASE_CURRENCY, FX_ROUNDING,
             f" +{FX_MARKUP_PCT:g}%" if FX_MARKUP_PCT else "")
    log.info("  FX source     : %s", _FX_CACHE.get("source") or "unavailable")
    for code, cfg in MARKETPLACES.items():
        cur   = cfg["currency"]
        price = market_price(cost, code)
        rate  = 1.0 if cur == BASE_CURRENCY else rates.get(cur)
        log.info("     %-2s %-3s  rate %-10s  sell %.2f %s -> %s",
                 code, cur, f"{rate:.4f}" if rate else "n/a", b["sell_raw"], PRICING_CURRENCY,
                 f"{price:,.2f} {cur}" if price is not None else "SKIPPED (no rate)")


# ── Google Sheet column indices — 0-based, must match the sheet header order ─
COL_SKU            = 0   # A — child SKU
COL_PARENT_SKU     = 1   # B — parent SKU (shared across variation siblings)
COL_PARENT_TITLE   = 2   # C — title for the non-purchasable parent listing
COL_TITLE          = 3   # D — full child listing title
COL_VARIATION      = 4   # E — variation label, e.g. "3GB / 15 Days"
COL_BRAND          = 5   # F — brand name, e.g. "YouGuide"
COL_PRICE          = 6   # G — WHOLESALE COST in BASE_CURRENCY (USD), no $ symbol.
                         #     NOT the shelf price: written by fetch_esim_prices.py
                         #     from the eSIM Access API, then marked up by
                         #     retail_price() (tier + min-profit, in GBP) and
                         #     converted per marketplace at the live FX rate.
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
# ── User input (continued) ──────────────────────────────────────────────────
COL_PRODUCT_ID     = 20  # U — user-supplied product identifier (EAN/GTIN barcode).
                         #     When set, it is listed on Amazon as the external
                         #     product ID (type EAN); when blank, the GTIN
                         #     exemption is claimed instead.
COL_MARKETPLACE    = 21  # V — COUNTRIES for this product (user input). One row
                         #     is published to EVERY country listed here: a single
                         #     code (UK), a comma list (UK,DE,FR), a region group
                         #     (EU/NA/FE) or ALL. Blank ⇒ UK only. No extra rows
                         #     are ever created — the one row drives all markets.
COL_PACKAGE_CODE   = 23  # X — eSIM Access packageCode this row is priced from.
                         #     fetch_esim_prices.py writes the code it matched by
                         #     NAME; fill it in by hand to pin a row to a specific
                         #     package (a manual entry is never overwritten).
COL_MARKET_STATE   = 22  # W — written by this script: compact JSON of per-market
                         #     state, e.g. {"UK":{"asin":"B0..","hash":".."}, ...}.
                         #     Lets the script skip markets already live at the
                         #     current content (no re-push). Do not edit by hand.

# Header labels for columns the script owns — written to row 1 if missing so a
# fresh sheet self-documents. Index → label. (product_id + marketplace are user
# input but are listed here so the columns are created/labelled on a fresh sheet.)
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
    COL_PRODUCT_ID:     "product_id",
    COL_MARKETPLACE:    "countries",
    COL_MARKET_STATE:   "market_state",
    COL_PACKAGE_CODE:   "package_code",
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


def sheet_content_hash(product: dict, price: Optional[float] = None) -> str:
    """Hash of EVERY user-editable content field — drives Sheet→Amazon updates.

    Any edit to title, price, quantity, description, image, etc. changes this
    hash and so triggers a push to Amazon (requirement: any column edit syncs).

    `price` overrides the sheet's price for the hash. Per-market callers pass the
    CONVERTED local price (see market_content_hash), so each market's stored hash
    tracks the number actually sent to that marketplace: an exchange-rate move
    that changes the shelf price in Poland re-pushes Poland only, while the
    base-currency (euro) markets hash exactly as they did before conversion
    existed.
    """
    payload = json.dumps(
        {
            "title":        product["title"],
            "parent_title": product["parent_title"],
            "variation":    product["variation"],
            "brand":        product["brand"],
            "price":        round(float(product["price"] if price is None else price), 2),
            "quantity":     int(product["quantity"]),
            "description":  product["description"],
            "image1":       product["image1"],
            "product_id":   (product.get("product_id") or "").strip(),
        },
        sort_keys=True,
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()[:16]


def market_content_hash(product: dict, code: str) -> Optional[str]:
    """The column-W hash for one marketplace — content + that market's own price.

    None when no exchange rate is available for the market's currency, which the
    callers treat as "cannot be pushed right now" (never as "unchanged").
    """
    price = market_price(product["price"], code)
    return None if price is None else sheet_content_hash(product, price)


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
    """Default (EU/UK) client — used by helper scripts and as the EU client."""
    return ListingsItems(
        credentials=SP_API_CREDENTIALS,
        marketplace=MARKETPLACE,
    )


# One ListingsItems client per token group, created on first use and cached. A
# None entry records "no token configured" so we don't retry building it.
_TOKEN_CLIENTS: dict = {}


def get_client_for_marketplace(code: str) -> Optional[ListingsItems]:
    """Return the cached client for a marketplace, or None if its token is unset.

    The client uses the marketplace's authorization-group refresh token
    (cfg["token"] → REFRESH_TOKENS) plus an endpoint marketplace from
    REGION_ENDPOINT (cfg["region"]) to resolve the host. JP and AU share the FE
    endpoint but use different tokens, so they get distinct clients.
    """
    cfg         = MARKETPLACES[code]
    token_group = cfg["token"]
    if token_group in _TOKEN_CLIENTS:
        return _TOKEN_CLIENTS[token_group]
    token = REFRESH_TOKENS.get(token_group, "")
    if not token:
        _TOKEN_CLIENTS[token_group] = None
        return None
    client = ListingsItems(
        credentials={"refresh_token": token, **LWA_CREDENTIALS},
        marketplace=REGION_ENDPOINT[cfg["region"]],
    )
    _TOKEN_CLIENTS[token_group] = client
    return client


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
    product_id:  str = "",
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
        # Under FBA the quantity is owned by Amazon and derives from units received
        # at a fulfilment centre, so the sheet's quantity is sent for FBM only.
        "fulfillment_availability": [
            {
                "fulfillment_channel_code": FULFILLMENT_CHANNEL,
                "marketplace_id":           MARKETPLACE_ID,
                **({} if USE_FBA else {"quantity": quantity}),
            }
        ],
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

    # Product identifier (col U). When a barcode is supplied, list it as the
    # external product ID (EAN); otherwise claim the GTIN exemption. Amazon
    # rejects providing both at once, so it is one or the other per child.
    pid = (product_id or "").strip()
    if pid:
        attributes["externally_assigned_product_identifier"] = [
            {"type": PRODUCT_ID_TYPE, "value": pid, "marketplace_id": MARKETPLACE_ID}
        ]
    else:
        attributes["supplier_declared_has_product_identifier_exemption"] = enum_attr(GTIN_EXEMPTION)

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
            "product_id":     cell(COL_PRODUCT_ID),
            "marketplace":    (cell(COL_MARKETPLACE) or DEFAULT_MARKETPLACE_CODE).upper(),
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
    price:           Optional[float] = None,
) -> tuple[bool, str]:
    """Create/update the parent (once per iteration) and the child on Amazon.

    `price` is the active marketplace's local price (already converted from the
    sheet's base currency). Omitted → the sheet's number is sent unchanged, which
    is only correct for a base-currency marketplace.
    """
    parent_sku = product["parent_sku"]
    # Key the "already created this pass" set by marketplace too: the same parent
    # SKU listed in several marketplaces must be created once in EACH of them.
    parent_key = (MARKETPLACE_ID, parent_sku)
    if parent_key not in created_parents:
        parent_payload = build_parent_payload(
            parent_title=product["parent_title"] or product["title"],
            brand=product["brand"],
            description=product["description"],
        )
        create_parent_listing(client, parent_sku, parent_payload)
        created_parents.add(parent_key)
        time.sleep(SLEEP_AFTER_PARENT)

    child_payload = build_child_payload(
        parent_sku=parent_sku,
        title=product["title"],
        brand=product["brand"],
        price=product["price"] if price is None else price,
        quantity=product["quantity"],
        description=product["description"],
        variation=product["variation"],
        image_url=product["image1"],
        product_id=product.get("product_id", ""),
    )
    return create_child_listing(client, product["sku"], child_payload)


def pull_updates(amazon_product: dict, code: str = HOME_MARKETPLACE) -> dict:
    """Build the column→value map that mirrors Amazon's values into the sheet.

    Empty values are skipped so a partial Amazon read never blanks a good cell.
    Price is skipped when 0 (almost always a read miss, not a genuine free item);
    quantity is always written because 0 legitimately means out of stock.

    PRICE IS NEVER PULLED BACK while TIER_PRICING is on: column G holds the base
    (cost) price that the tier formula reads, so mirroring Amazon's shelf price
    into it would feed a marked-up, FX-converted number back through the markup on
    the next pass. With TIER_PRICING off it is pulled, converted from `code`'s
    currency back to BASE_CURRENCY (and skipped when that rate is unknown).
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
    if float(p.get("price") or 0) > 0 and not TIER_PRICING:
        in_base = base_price_from(float(p["price"]), code)
        if in_base is not None:
            updates[COL_PRICE] = in_base
        else:
            log.warning("Reverse sync: no %s rate — price left unchanged for %s",
                        MARKETPLACES.get(code, {}).get("currency", "?"), p.get("sku", ""))
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
) -> tuple[bool, str]:
    """Write the outcome of a push back to the sheet, verifying first.

    Returns (ok, status):
      * ok=True  → Amazon accepted it and verification found no ERROR issues; the
                   row is marked live (UPLOADED/UPDATED + hash snapshot stored).
      * ok=False → the API accepted the submission but a follow-up verify found an
                   ERROR (e.g. a deleted/limbo SKU Amazon won't put in the
                   catalogue). The row is marked RETRY and left NOT-live, so the
                   next poll re-attempts it until Amazon finally accepts.
    """
    sku  = product["sku"]
    asin = product["asin"] or fetch_asin(client, sku)

    if VERIFY_AFTER_SUBMIT:
        v_status, v_issues = verify_listing(client, sku)
        if v_issues:
            # Accepted by the API but rejected during processing → not live yet.
            # Keep the row retryable (status not DELETED/HOLD/UPLOADED/UPDATED).
            status = f"RETRY ⚠ {v_issues}"[:400]
            write_cells(worksheet, sheet_row, {
                COL_UPLOADED:    "",
                COL_STATUS:      status,
                COL_ASIN:        asin or "",
                COL_LAST_SYNCED: now_iso(),
                COL_CONFLICT:    conflict_note,
            })
            log.warning("Row %d (%s): submission accepted but NOT live — will retry. %s",
                        sheet_row, sku, v_issues)
            return False, status
        status = f"{base_status} ({v_status})" if v_status else base_status
    else:
        status = base_status

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
    return True, status


# =============================================================================
#  MAIN ORCHESTRATION
# =============================================================================

def load_market_state(cell: str) -> dict:
    """Parse the per-market state JSON from column W into a dict (or {})."""
    cell = (cell or "").strip()
    if not cell:
        return {}
    try:
        data = json.loads(cell)
        return data if isinstance(data, dict) else {}
    except (ValueError, TypeError):
        return {}


def dump_market_state(state: dict) -> str:
    """Serialise per-market state to compact JSON for column W."""
    try:
        return json.dumps(state, separators=(",", ":"), sort_keys=True)
    except (ValueError, TypeError):
        return ""


def summarize_results(results: dict) -> str:
    """Build a human-readable per-market status summary for column L.

    e.g. {"UK":"OK","DE":"OK","JP":"RETRY"} → "OK:DE,UK | RETRY:JP".
    """
    groups: dict = {}
    for code, outcome in results.items():
        groups.setdefault(outcome, []).append(code)
    ordered = ["OK", "RETRY", "FAIL", "NO_TOKEN", "NO_SELLER", "NO_FX", "NO_COST", "BAD_CODE"]
    parts = []
    for key in ordered + [k for k in groups if k not in ordered]:
        if key in groups:
            parts.append(f"{key}:{','.join(sorted(groups[key]))}")
    return " | ".join(parts)


def preview_plan(data_rows: list) -> None:
    """Log a one-line-per-row summary of what the script intends to do.

    This is a heads-up printed BEFORE any Amazon reads, derived purely from the
    sheet's Status (L), Action (O) and hash (N) columns. The live Amazon read in
    the main pass can still refine a CREATE/UPDATE/no-change row into PULL
    (Amazon was edited externally) or "DELETED (on Amazon)" (removed on Amazon),
    so treat this as the *intended* plan, not a guarantee.
    """
    log.info("%s", "-" * 76)
    log.info("PLAN for this pass (one row per product; pushed to every country in col V):")
    log.info("  %-4s | %-10s | %-20s | %-16s | %-5s | %s",
             "row", "COUNTRIES", "SKU", "STATUS", "ACT", "DECISION")
    log.info("  %s", "-" * 72)

    counts: dict = {}
    for row_index, raw_row in enumerate(data_rows):
        if not any(c.strip() for c in raw_row):
            continue
        sheet_row = row_index + 2

        def raw_cell(col: int) -> str:
            return str(raw_row[col]).strip() if col < len(raw_row) else ""

        sku          = raw_cell(COL_SKU)
        status_raw   = raw_cell(COL_STATUS)
        action_raw   = raw_cell(COL_ACTION)
        countries    = raw_cell(COL_MARKETPLACE) or HOME_MARKETPLACE
        status       = status_raw.upper()
        want_deleted = action_raw.upper() in DELETE_TRIGGERS

        codes, _ = parse_countries(countries)
        if not codes:
            codes = [HOME_MARKETPLACE]

        if not sku:
            decision = "skip (no SKU)"
        elif want_deleted:
            decision = f"DELETE ({len(codes)} market[s])"
        elif status == "HOLD":
            decision = "HOLD (frozen)"
        else:
            product = parse_row(raw_row, row_index)
            if product is None:
                decision = "skip (missing required fields)"
            else:
                state = load_market_state(raw_cell(COL_MARKET_STATE))
                # Each market's hash covers its own converted price, so a market
                # whose local price moved with the exchange rate counts as needing
                # a push here (a market with no rate available also counts, and is
                # reported NO_FX by the live pass).
                live  = sum(1 for c in codes
                            if (state.get(c) or {}).get("asin")
                            and (state.get(c) or {}).get("hash") == market_content_hash(product, c))
                need  = len(codes) - live
                decision = f"PUSH {need}/{len(codes)}" if need else f"all live ({len(codes)})"

        counts[decision.split(" ")[0]] = counts.get(decision.split(" ")[0], 0) + 1
        log.info("  %-4d | %-10s | %-20s | %-16s | %-5s | %s",
                 sheet_row, countries[:10], (sku or "—")[:20], (status_raw or "—")[:16],
                 (action_raw or "—")[:5], decision)

    summary = ", ".join(f"{v}× {k}" for k, v in sorted(counts.items()))
    log.info("  %s", "-" * 72)
    log.info("PLAN summary: %s", summary or "nothing to do")
    log.info("%s", "-" * 76)


def process_sheet(
    worksheet: gspread.Worksheet,
    iteration: int,
) -> int:
    """One full pass. Returns the number of rows changed (used to drive the
    activity-based poll interval).

    One sheet row = one product. It is pushed to EVERY marketplace listed in its
    countries column (V); no extra rows are created. Per-market state is kept in
    column W so unchanged markets are skipped without an API call."""
    all_values = worksheet.get_all_values()

    if len(all_values) < 2:
        log.info("Iteration %d: sheet has no data rows — nothing to do", iteration)
        return 0

    data_rows = all_values[1:]  # strip header row
    log.info("Iteration %d: loaded %d data rows", iteration, len(data_rows))

    preview_plan(data_rows)

    created_parents: set = set()        # (marketplace_id, parent_sku) pushed this pass
    processed_skus:  dict = {}          # sku -> sheet_row of first occurrence
    pushed = deleted = failed = skipped = 0

    # --first <text> moves matching rows to the FRONT of the pass. Nothing is
    # excluded — the rest follow in sheet order — so it only changes WHEN a row is
    # reached, which is how you get one product live and verifiable before a
    # catalogue-wide repricing works through the remaining hundreds of rows.
    # The original sheet position travels with each row, so sheet_row stays right.
    ordered = list(enumerate(data_rows))
    if PRIORITY_FIRST:
        needle = PRIORITY_FIRST.lower()

        def _is_priority(pair) -> bool:
            row = pair[1]
            hay = " ".join(str(row[c]) for c in (COL_SKU, COL_TITLE, COL_PARENT_TITLE)
                           if c < len(row)).lower()
            return needle in hay

        head = [p for p in ordered if _is_priority(p)]
        tail = [p for p in ordered if not _is_priority(p)]
        ordered = head + tail
        log.info("Priority: %d row(s) matching %r go first, then the other %d",
                 len(head), PRIORITY_FIRST, len(tail))

    for row_index, raw_row in ordered:
        if not any(c.strip() for c in raw_row):
            continue   # blank row

        sheet_row = row_index + 2   # 1-based sheet row (row 2 = first data row)

        def raw_cell(col: int) -> str:
            return str(raw_row[col]).strip() if col < len(raw_row) else ""

        action_raw     = raw_cell(COL_ACTION).upper()
        current_status = raw_cell(COL_STATUS).upper()
        sku            = raw_cell(COL_SKU)
        want_deleted   = action_raw in DELETE_TRIGGERS

        if not sku:
            continue   # nothing to act on without a SKU

        # -- HOLD: frozen row, do nothing --------------------------------------
        if current_status == "HOLD":
            log.info("Row %d (%s): on hold -- skipping", sheet_row, sku)
            skipped += 1
            continue

        # Resolve the target marketplaces from the countries column (V).
        codes, invalid = parse_countries(raw_cell(COL_MARKETPLACE))
        if not codes:
            codes = [HOME_MARKETPLACE]            # blank => UK only
        results: dict = {c: "BAD_CODE" for c in invalid}
        if invalid:
            log.warning("Row %d (%s): unknown country code(s) %s",
                        sheet_row, sku, ", ".join(invalid))

        # Guard against the same SKU appearing on more than one row.
        if sku in processed_skus:
            msg = f"SKIPPED: duplicate of row {processed_skus[sku]} (same SKU)"
            log.warning("Row %d (%s): %s", sheet_row, sku, msg)
            write_cells(worksheet, sheet_row, {COL_STATUS: msg})
            skipped += 1
            continue
        processed_skus[sku] = sheet_row

        state = load_market_state(raw_cell(COL_MARKET_STATE))

        # -- DELETE requested (action column O) -> remove from every market ----
        if want_deleted:
            if not state and current_status.startswith("DELETED"):
                skipped += 1
                continue
            targets = [c for c in sorted(set(list(state.keys()) + codes)) if c in MARKETPLACES]
            for code in targets:
                client = get_client_for_marketplace(code)
                if client is None:
                    results[code] = "NO_TOKEN"
                    continue
                set_active_marketplace(code)
                if not SELLER_ID:
                    results[code] = "NO_SELLER"
                    continue
                ok, _ = delete_listing(client, sku)
                results[code] = "DELETED" if ok else "FAIL"
                if ok:
                    state.pop(code, None)
                time.sleep(SLEEP_AFTER_CHILD)
            write_cells(worksheet, sheet_row, {
                COL_UPLOADED:     "",
                COL_STATUS:       ("DELETED | " + summarize_results(results)) if results else "DELETED",
                COL_MARKET_STATE: dump_market_state(state),
                COL_LAST_SYNCED:  now_iso(),
            })
            deleted += 1
            continue

        # -- Build the product and push to every listed marketplace ------------
        product = parse_row(raw_row, row_index)
        if product is None:
            failed += 1
            continue
        sheet_hash = sheet_content_hash(product)

        any_push = False
        for code in codes:
            client = get_client_for_marketplace(code)
            if client is None:
                results[code] = "NO_TOKEN"
                continue
            set_active_marketplace(code)
            if not SELLER_ID:
                results[code] = "NO_SELLER"
                continue
            prev = state.get(code) or {}
            # Convert the sheet's base-currency price into this marketplace's own
            # currency at today's rate. No rate → skip the market entirely; sending
            # the unconverted number would list e.g. 29.99 EUR as 29.99 PLN.
            local_price = market_price(product["price"], code)
            if local_price is None:
                results[code] = "NO_FX"
                log.warning("Row %d (%s): no %s exchange rate — %s skipped this pass",
                            sheet_row, sku, MARKETPLACES[code]["currency"], code)
                continue
            if local_price <= 0:
                results[code] = "NO_COST"
                log.warning("Row %d (%s): no cost price in column G — %s skipped "
                            "(run fetch_esim_prices.py)", sheet_row, sku, code)
                continue
            market_hash = sheet_content_hash(product, local_price)
            # Skip markets already live at the current content AND local price.
            if prev.get("asin") and prev.get("hash") == market_hash:
                results[code] = "OK"
                continue

            any_push = True
            sell_gbp = retail_price_raw(product["price"])
            log.info("Row %d (%s) %s: cost %.2f %s -> sell %.2f %s -> %s %s",
                     sheet_row, sku, code, product["price"], BASE_CURRENCY,
                     sell_gbp or 0.0, PRICING_CURRENCY,
                     f"{local_price:,.2f}", CURRENCY)
            success, status_msg = push_to_amazon(client, product, created_parents,
                                                 price=local_price)
            if not success:
                results[code] = "FAIL"
                state[code] = {"asin": prev.get("asin", ""), "hash": "", "err": status_msg[:120]}
                time.sleep(SLEEP_AFTER_CHILD)
                continue

            if VERIFY_AFTER_SUBMIT:
                v_status, v_issues = verify_listing(client, sku)
                if v_issues:
                    results[code] = "RETRY"
                    state[code] = {"asin": prev.get("asin", ""), "hash": "", "err": v_issues[:120]}
                    time.sleep(SLEEP_AFTER_CHILD)
                    continue

            asin = prev.get("asin") or fetch_asin(client, sku)
            results[code] = "OK"
            # "price"/"cur" are informational — they make column W readable at a
            # glance ("what is this SKU actually selling for in Japan?").
            state[code] = {"asin": asin or "", "hash": market_hash,
                           "price": local_price, "cur": CURRENCY}
            time.sleep(SLEEP_AFTER_CHILD)

        # -- Write the per-market outcome back to the single row ---------------
        home_asin = (state.get(HOME_MARKETPLACE) or {}).get("asin", "") \
            or next((v.get("asin", "") for v in state.values() if v.get("asin")), "")
        ok_any  = any(r == "OK" for r in results.values())
        bad_any = any(r in ("FAIL", "RETRY") for r in results.values())
        updates = {
            COL_UPLOADED:     "YES" if ok_any else "",
            COL_STATUS:       summarize_results(results),
            COL_ASIN:         home_asin,
            COL_HASH:         sheet_hash,
            COL_MARKET_STATE: dump_market_state(state),
            COL_LAST_SYNCED:  now_iso(),
        }
        if any_push:
            updates[COL_SHEET_UPDATED] = now_iso()
        write_cells(worksheet, sheet_row, updates)

        if any_push:
            pushed += 1
        else:
            skipped += 1
        if bad_any:
            failed += 1

    log.info("%s", "-" * 76)
    log.info(
        "ITERATION %d DONE  |  pushed:%d  deleted:%d  failed:%d  skipped:%d",
        iteration, pushed, deleted, failed, skipped,
    )
    return pushed + deleted


def run() -> None:
    log.info("%s", "=" * 76)
    log.info("Amazon SP-API two-way sync%s", "   *** DRY_RUN (no writes) ***" if DRY_RUN else "   *** LIVE — writing to Amazon ***")
    log.info("%s", "=" * 76)
    # Group the configured marketplaces by token authorization group and show
    # which are usable (group has a refresh token) vs skipped (token missing).
    for group in ("EU", "NA", "JP", "AU"):
        codes = [c for c, cfg in MARKETPLACES.items() if cfg["token"] == group]
        if not codes:
            continue
        have_tok = bool(REFRESH_TOKENS.get(group))
        have_sid = bool(SELLER_IDS.get(group))
        flags = []
        if not have_tok:
            flags.append("set SP_API_REFRESH_TOKEN_%s" % group)
        if not have_sid:
            flags.append("set SELLER_ID_%s" % group)
        log.info("  %-2s token:%s seller:%s : %s%s", group,
                 "OK" if have_tok else "--", "OK" if have_sid else "--",
                 ", ".join(codes),
                 ("   (skipped — " + "; ".join(flags) + ")") if flags else "")
    log.info("  English-only   : %s", ENGLISH_ONLY)
    log.info("  Default (blank): %s", DEFAULT_MARKETPLACE_CODE)
    log.info("  Product type   : %s", PRODUCT_TYPE)
    log.info("  Poll cadence   : %d–%ds (activity-based)", POLL_MIN_SECONDS, POLL_MAX_SECONDS)
    log.info("  Delete triggers: %s", sorted(DELETE_TRIGGERS))
    log.info("  Google Sheet   : %s / %r", SPREADSHEET_ID, WORKSHEET_NAME)
    # Fetch the rates once up front so the table below (and the first pass) use
    # real numbers, and any feed problem is visible before anything is pushed.
    log_fx_table()
    log.info("%s", "=" * 76)

    worksheet = get_worksheet()
    ensure_headers(worksheet)

    iteration = 0
    interval  = POLL_MIN_SECONDS
    while True:
        iteration += 1
        log.info(" ")
        log.info("%s", "=" * 76)
        log.info("ITERATION %d   started %s", iteration, now_iso())
        log.info("%s", "=" * 76)
        try:
            # Refresh FX before the pass; fx_rates() is a no-op inside its TTL, so
            # this only actually fetches every FX_TTL_HOURS.
            fx_rates()
            changes = process_sheet(worksheet, iteration)
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

        mins = interval / 60
        log.info("%s", "-" * 76)
        log.info("Sleeping %ds (%.0f min) until next poll…  (Ctrl-C to stop)", interval, mins)
        log.info(" ")
        try:
            time.sleep(interval)
        except KeyboardInterrupt:
            log.info("Interrupted by user — exiting")
            return


if __name__ == "__main__":
    # --rates [PRICE] — read-only FX check: print today's rates and the shelf
    # price each marketplace would get. Touches neither Amazon nor the sheet.
    if "--rates" in sys.argv:
        args = [a for a in sys.argv[1:] if a != "--rates"]
        try:
            sample = float(args[0]) if args else 29.99
        except ValueError:
            sample = 29.99
        log.info("Pricing check — cost %.2f %s through tiers + FX, in every marketplace",
                 sample, BASE_CURRENCY)
        log_fx_table(sample)
        sys.exit(0)

    # --first <text> — process rows whose SKU/title contain <text> before the rest
    # (e.g. --first slovakia to watch one product land before the catalogue does).
    if "--first" in sys.argv:
        pos = sys.argv.index("--first")
        if pos + 1 < len(sys.argv):
            PRIORITY_FIRST = sys.argv[pos + 1]
            log.info("Rows matching %r will be processed first", PRIORITY_FIRST)
        else:
            log.error("--first needs a value, e.g. --first slovakia")
            sys.exit(2)

    print("Script Started")
    print("Verifing environment variables...")
    run()