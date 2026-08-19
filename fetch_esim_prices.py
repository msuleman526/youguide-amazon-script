#!/usr/bin/env python3
"""
eSIM Access → Google Sheet: write each row's real WHOLESALE COST (USD)
=====================================================================
Fills column G (the cost the pricing formula reads) from the live eSIM Access
catalogue, so the sheet stops carrying hand-typed prices.

    packages = POST https://api.esimaccess.com/api/v1/open/package/list
    row      = matched to a package BY NAME
    column G = package price ÷ 10000        (the API quotes 1/10000 USD)
    column X = the packageCode it matched   (audit trail + manual pin)

upload_listings.py then does the rest on its next poll: tier markup in GBP
(cost ×3/×2/×1.6/×1.3 + £4 minimum profit after ~18% Amazon fees, rounded up to
the next .99) and conversion into each marketplace's currency.

MATCHING IS BY NAME, not SKU. Both sides are reduced to (place, GB, days):
    sheet  "Bulgaria eSIM 3GB 15Days | Pay As You Go"  → ("bulgaria", 3.0, 15)
    API    "Bulgaria 3GB 15Days"                       → ("bulgaria", 3.0, 15)
Marketing words (eSIM, Travel Data SIM, Prepaid, YouGuide…), anything after a
"|", and area counts ("Europe (30+ areas)") are stripped first. A row whose key
matches nothing — or matches several packages at different prices — is REPORTED,
never guessed at: put the packageCode in column X and it is used verbatim.

Usage:
    python fetch_esim_prices.py                 # preview: match + price report, writes NOTHING
    python fetch_esim_prices.py --apply         # write columns G and X to the sheet
    python fetch_esim_prices.py --find bulgaria # search the catalogue by name
    python fetch_esim_prices.py --refresh       # ignore the cached catalogue

Credentials live in .env (never in source):
    ESIM_ACCESS_CODE=...      # "Access Code" from console.esimaccess.com
    ESIM_SECRET_KEY=...       # only needed if the account requires signing
"""

import json
import os
import re
import sys
import time
import unicodedata
import urllib.error
import urllib.request

import gspread

from upload_listings import (
    BASE_CURRENCY,
    COL_PACKAGE_CODE,
    COL_PARENT_TITLE,
    COL_PRICE,
    COL_SKU,
    COL_TITLE,
    COL_VARIATION,
    PRICING_CURRENCY,
    ensure_headers,
    get_worksheet,
    log,
    retail_price,
)

# ── API ──────────────────────────────────────────────────────────────────────
API_URL      = "https://api.esimaccess.com/api/v1/open/package/list"
ACCESS_CODE  = os.getenv("ESIM_ACCESS_CODE", "")
SECRET_KEY   = os.getenv("ESIM_SECRET_KEY", "")
API_TIMEOUT  = 90

# The API quotes money in 1/10000 of a unit: 112500 → 11.25.
PRICE_DIVISOR = 10000.0

# Every package carries its own `currencyCode`, and it is whatever the eSIM Access
# ACCOUNT is denominated in — the /package/list request takes no currency
# parameter, so the supplier cannot be asked for euros; the account has to be
# switched. Today every one of the ~2,900 packages says USD. Column G is read as
# BASE_CURRENCY, so if the account is ever moved to EUR, set BASE_CURRENCY=EUR in
# .env and the check below goes quiet — without it, euro costs would be priced as
# though they were dollars and every listing would be silently ~15% cheap.

# Catalogue cache — 2,900 packages is a slow call, and re-running the matcher
# while tuning it shouldn't hammer the supplier.
CACHE_FILE     = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              "esim_packages.json")
CACHE_TTL_SECS = 6 * 3600


# ── Name normalisation ───────────────────────────────────────────────────────
# Words that describe the product rather than identify the package. Stripped from
# both sides before matching so "Zambia eSIM 3GB" and "Zambia 3GB" agree.
NOISE_WORDS = [
    "esim", "e-sim", "sim card", "sim", "travel data", "data card", "data",
    "prepaid", "pay as you go", "payg", "unlimited plan", "package", "plan",
    "youguide", "card", "mobile", "internet", "roaming", "tourist",
]

# Place-name spellings that differ between the sheet and the supplier.
PLACE_ALIASES = {
    "aaland islands": "aland islands",   # sheet spells Åland with a double A
    "uk": "united kingdom",
    "usa": "united states",
    "us": "united states",
    "uae": "united arab emirates",
    "holland": "netherlands",
    "south korea": "korea",
    "hongkong": "hong kong",
}


def _clean(text: str) -> str:
    """Lowercase, drop anything after a '|', collapse punctuation/whitespace.

    Accents are folded first (NFKD, combining marks dropped) so the supplier's
    "Curaçao"/"Åland Islands" reduce to the sheet's plain-ASCII spelling instead
    of losing the accented letter entirely.
    """
    t = unicodedata.normalize("NFKD", str(text or ""))
    t = "".join(ch for ch in t if not unicodedata.combining(ch))
    t = t.split("|")[0].lower()
    t = t.replace("&", " and ")
    t = re.sub(r"\(.*?\)", " ", t)          # "(30+ areas)"
    t = re.sub(r"[^a-z0-9./+ -]", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def parse_name(text: str) -> tuple:
    """Reduce a package/listing name to its identity: (place, gb, days).

    days is None for a per-day plan ("Singapore 1GB/Day"), which is a different
    product from a fixed bundle and must never match one.
    """
    t = _clean(text)

    vol_match = re.search(r"(\d+(?:\.\d+)?)\s*gb", t)
    gb        = float(vol_match.group(1)) if vol_match else None

    daily = bool(re.search(r"gb\s*/\s*day|/\s*day\b|\bdaily\b", t))
    days  = None
    if not daily:
        day_match = re.search(r"(\d+)\s*days?\b", t)
        if day_match:
            days = int(day_match.group(1))

    # Place = everything before the volume, minus the marketing vocabulary.
    place = t[:vol_match.start()] if vol_match else t
    for w in NOISE_WORDS:
        place = place.replace(w, " ")
    place = re.sub(r"-\d+\b", " ", place)        # "asia-20" → "asia"
    place = re.sub(r"\s+", " ", place).strip(" -")
    place = PLACE_ALIASES.get(place, place)

    return (place, gb, days if not daily else "daily")


# ── Catalogue ────────────────────────────────────────────────────────────────
def fetch_packages(refresh: bool = False) -> list:
    """The supplier's full package list, cached on disk for CACHE_TTL_SECS."""
    if not refresh and os.path.exists(CACHE_FILE):
        age = time.time() - os.path.getmtime(CACHE_FILE)
        if age < CACHE_TTL_SECS:
            try:
                with open(CACHE_FILE, encoding="utf-8") as fh:
                    packages = json.load(fh)
                log.info("Catalogue: %d packages from cache (%.1f h old)",
                         len(packages), age / 3600)
                return packages
            except (OSError, ValueError):
                pass   # unreadable cache → fetch fresh

    if not ACCESS_CODE:
        log.error("ESIM_ACCESS_CODE is not set in .env — cannot call the API")
        return []

    body    = json.dumps({"locationCode": "", "type": "", "packageCode": "",
                          "iccid": ""}).encode("utf-8")
    headers = {"Content-Type": "application/json", "RT-AccessCode": ACCESS_CODE}
    if SECRET_KEY:
        # Some accounts require the secret alongside the access code; sending it
        # is harmless where it is ignored.
        headers["RT-SecretKey"] = SECRET_KEY

    try:
        req  = urllib.request.Request(API_URL, data=body, headers=headers)
        with urllib.request.urlopen(req, timeout=API_TIMEOUT) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except (urllib.error.URLError, OSError, ValueError) as exc:
        log.error("eSIM Access API call failed: %s", exc)
        return []

    if not data.get("success"):
        log.error("eSIM Access API error %s: %s", data.get("errorCode"), data.get("errorMsg"))
        return []

    packages = (data.get("obj") or {}).get("packageList") or []
    log.info("Catalogue: %d packages fetched from the API", len(packages))
    try:
        with open(CACHE_FILE, "w", encoding="utf-8") as fh:
            json.dump(packages, fh)
    except OSError as exc:
        log.warning("Could not cache the catalogue: %s", exc)
    return packages


def cost_usd(package: dict) -> float:
    """A package's wholesale cost in whole currency units (API quotes 1/10000)."""
    try:
        return round(float(package.get("price") or 0) / PRICE_DIVISOR, 2)
    except (TypeError, ValueError):
        return 0.0


def check_currency(packages: list) -> None:
    """Warn loudly if the catalogue is not quoted in BASE_CURRENCY.

    A mismatch is not a crash — the run still shows what it would do — but the
    numbers would be wrong, so it must never pass unnoticed.
    """
    found = {str(p.get("currencyCode") or "?").upper() for p in packages if p.get("price")}
    if not found or found == {BASE_CURRENCY}:
        return
    log.warning("!! CATALOGUE CURRENCY MISMATCH: supplier quotes %s, but column G "
                "is read as BASE_CURRENCY=%s", "/".join(sorted(found)), BASE_CURRENCY)
    log.warning("   Set BASE_CURRENCY=%s in .env (this reprices every listing) or "
                "switch the eSIM Access account back to %s.",
                sorted(found)[0], BASE_CURRENCY)


def build_index(packages: list) -> tuple:
    """(index by name-key, index by packageCode). Cheapest option kept per key."""
    by_key: dict = {}
    by_code: dict = {}
    for p in packages:
        by_code[str(p.get("packageCode") or "")] = p
        key = parse_name(p.get("name", ""))
        if key[1] is None:            # no data volume in the name → unmatchable
            continue
        by_key.setdefault(key, []).append(p)
    for key, options in by_key.items():
        options.sort(key=cost_usd)    # cheapest first — that is the cost we quote
    return by_key, by_code


# ── Matching ─────────────────────────────────────────────────────────────────
def row_key(raw_row: list) -> tuple:
    """The (place, gb, days) key for a sheet row, read from its NAME columns.

    The title carries the place; the variation cell ("3GB / 15 Days") is the more
    reliable source of volume/duration, so it fills in whatever the title lacks.
    """
    def cell(col: int) -> str:
        return str(raw_row[col]).strip() if col < len(raw_row) else ""

    title = cell(COL_TITLE) or cell(COL_PARENT_TITLE)
    place, gb, days = parse_name(title)
    if gb is None or days is None:
        _, v_gb, v_days = parse_name(f"x {cell(COL_VARIATION)}")
        gb   = gb if gb is not None else v_gb
        days = days if days is not None else v_days
    return (place, gb, days)


def match_row(raw_row: list, by_key: dict, by_code: dict) -> tuple:
    """→ (package|None, note). A manual packageCode in column X always wins."""
    def cell(col: int) -> str:
        return str(raw_row[col]).strip() if col < len(raw_row) else ""

    # A pin is one packageCode — short and unspaced. Anything else in column X is
    # stray text (the sheet has some), so fall through to name matching instead of
    # reporting a nonsense "code not in catalogue".
    pinned = cell(COL_PACKAGE_CODE)
    if pinned and len(pinned) <= 24 and re.fullmatch(r"[A-Za-z0-9_-]+", pinned):
        pkg = by_code.get(pinned)
        return (pkg, "pinned" if pkg else f"PINNED CODE {pinned} NOT IN CATALOGUE")

    key = row_key(raw_row)
    if key[1] is None or key[2] is None:
        return (None, f"cannot read GB/days from the name {key}")

    options = by_key.get(key)
    if not options and key[2] == 1:
        # A 1-day bundle and a per-day plan are the same product from opposite
        # naming conventions: our "Zambia eSIM 1GB 1 Day" is the supplier's
        # "Zambia 1GB/Day". Only ever tried for exactly one day.
        options = by_key.get((key[0], key[1], "daily"))
        if options:
            return (options[0], "matched as a /Day plan")
    if not options:
        return (None, f"no package named like {key}")
    if len(options) > 1:
        spread = [f"{p['packageCode']}=${cost_usd(p):.2f}" for p in options[:4]]
        return (options[0], "cheapest of %d (%s)" % (len(options), ", ".join(spread)))
    return (options[0], "matched")


# ── Reporting / writing ──────────────────────────────────────────────────────
def run(apply_changes: bool, refresh: bool) -> int:
    packages = fetch_packages(refresh)
    if not packages:
        return 1
    check_currency(packages)
    by_key, by_code = build_index(packages)

    worksheet = get_worksheet()
    if apply_changes:
        ensure_headers(worksheet)
    rows = worksheet.get_all_values()
    if len(rows) < 2:
        log.info("Sheet has no data rows")
        return 0

    updates:  list = []
    unmatched: list = []
    log.info("%s", "-" * 110)
    log.info("%-4s %-22s %-34s %-9s %-8s %-10s %s",
             "row", "sku", "matched package", "code", "cost$", f"sell({PRICING_CURRENCY})", "note")
    log.info("%s", "-" * 110)

    for i, raw in enumerate(rows[1:]):
        if not any(str(c).strip() for c in raw):
            continue
        sheet_row = i + 2
        sku = str(raw[COL_SKU]).strip() if COL_SKU < len(raw) else ""
        if not sku:
            continue

        pkg, note = match_row(raw, by_key, by_code)
        if pkg is None:
            unmatched.append((sheet_row, sku, note))
            log.warning("%-4d %-22s %-34s %-9s %-8s %-10s %s",
                        sheet_row, sku[:22], "— NO MATCH —", "", "", "", note)
            continue

        cost = cost_usd(pkg)
        sell = retail_price(cost)
        old  = str(raw[COL_PRICE]).strip() if COL_PRICE < len(raw) else ""
        log.info("%-4d %-22s %-34s %-9s %-8s %-10s %s",
                 sheet_row, sku[:22], str(pkg.get("name", ""))[:34],
                 pkg.get("packageCode", ""), f"{cost:.2f}",
                 f"{sell:.2f}" if sell else "n/a", note)

        # Only write when something actually changes — keeps the Sheets quota low.
        cell_updates = {}
        if old != f"{cost:.2f}":
            cell_updates[COL_PRICE] = f"{cost:.2f}"
        # Write the code when the cell is empty OR holds something that isn't a
        # package code (the sheet has stray text in X). A REAL pin is never
        # touched — match_row() already returned that package.
        have_code = str(raw[COL_PACKAGE_CODE]).strip() if COL_PACKAGE_CODE < len(raw) else ""
        if have_code != pkg.get("packageCode", ""):
            cell_updates[COL_PACKAGE_CODE] = pkg.get("packageCode", "")
        if cell_updates:
            updates.append((sheet_row, cell_updates, sku, old, cost))

    log.info("%s", "-" * 110)
    log.info("%d rows would change, %d unmatched", len(updates), len(unmatched))
    if unmatched:
        log.info("Unmatched rows need a packageCode in column X (find one with "
                 "--find <name>):")
        for sheet_row, sku, note in unmatched:
            log.info("   row %-4d %-24s %s", sheet_row, sku, note)

    if not apply_changes:
        log.info("PREVIEW ONLY — nothing was written. Re-run with --apply to write "
                 "columns G (cost) and X (package code).")
        return 0

    cells = []
    for sheet_row, cell_updates, *_ in updates:
        for col, value in cell_updates.items():
            cells.append(gspread.Cell(sheet_row, col + 1, value))
    if cells:
        worksheet.update_cells(cells)
        log.info("WROTE %d cells across %d rows", len(cells), len(updates))
    else:
        log.info("Nothing to write — the sheet already matches the catalogue")
    return 0


def find(query: str) -> int:
    """Search the catalogue by name — for filling column X by hand."""
    packages = fetch_packages(False)
    q = _clean(query)
    hits = [p for p in packages if q in _clean(p.get("name", ""))]
    log.info("%d packages matching %r", len(hits), query)
    for p in sorted(hits, key=cost_usd)[:60]:
        log.info("   %-9s %-40s $%-8.2f %s",
                 p.get("packageCode", ""), str(p.get("name", ""))[:40],
                 cost_usd(p), p.get("slug", ""))
    return 0


if __name__ == "__main__":
    args = sys.argv[1:]
    if "--find" in args:
        idx = args.index("--find")
        sys.exit(find(" ".join(args[idx + 1:]) or ""))
    sys.exit(run(apply_changes="--apply" in args, refresh="--refresh" in args))
