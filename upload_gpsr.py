#!/usr/bin/env python3
"""
upload_gpsr.py — bulk-apply EU GPSR compliance to existing Amazon listings
==========================================================================
The EU General Product Safety Regulation (GPSR) requires every product sold in
the EU to carry, on its listing:
  * the MANUFACTURER's contact (email/URL),
  * the EU RESPONSIBLE PERSON's contact (email/URL),
  * a safety attestation, and
  * the product's safety / warning documentation.

Amazon exposes these as SIM_CARD listing attributes. This script PATCHES only
those GPSR attributes onto each SKU that already exists on Amazon — it does NOT
rebuild or re-price the listing (unlike upload_listings.py's full PUT). It
reuses upload_listings.py's credentials, per-region clients, marketplace
registry and the same Google Sheet, so there is nothing new to configure except
the values in the CONFIG block below.

  Amazon requires GPSR to be submitted ONCE PER MARKETPLACE (there is no
  "all of EU" upload), but bulk WITHIN a marketplace — which is exactly what
  this does: one PATCH per (SKU × marketplace), looped over every SKU in the
  sheet and every GPSR marketplace.

Safe by default: NOTHING is written unless you pass --live. Without it the
script prints the PATCH it WOULD send (a dry run).

Usage:
    python upload_gpsr.py                                   # dry run, ALL skus × ALL GPSR markets
    python upload_gpsr.py --sku HAJJ-LUX-10GB-30D --market DE          # dry run, one sku/market
    python upload_gpsr.py --sku HAJJ-LUX-10GB-30D --market DE --live   # LIVE: patch that one
    python upload_gpsr.py --market DE --live                # LIVE: whole DE marketplace
    python upload_gpsr.py --limit 5 --live                  # LIVE: first 5 skus (all markets)
    python upload_gpsr.py --live                            # LIVE: full bulk apply

Recommended first run (matches the agreed rollout):
    1. python upload_gpsr.py --sku <one-sku> --market DE            (dry run — inspect the payload)
    2. python upload_gpsr.py --sku <one-sku> --market DE --live     (patch one, then check the detail page)
    3. python upload_gpsr.py --live                                 (bulk, once step 2 looks right)
"""

import argparse
import sys
import time

from sp_api.base import SellingApiException

from upload_listings import (
    log,
    PRODUCT_TYPE,
    MARKETPLACES,
    SELLER_IDS,
    get_worksheet,
    get_client_for_marketplace,
    COL_SKU,
)

# =============================================================================
#  CONFIG  — the GPSR facts. Edit here; nothing else needs to change.
# =============================================================================

# Manufacturer & EU Responsible Person contact. GPSR listing attributes carry
# only an email/URL that REFERENCES the party; the full postal address
# (YouGuide International BV, MC Square, Leonardo Da Vincilaan 19, 1831 Diegem,
# Belgium) must be registered ONCE per marketplace in Seller Central →
# "Manage Your Compliance". The email below MUST match what you registered there.
MANUFACTURER_EMAIL       = "support@youguide.com"
RESPONSIBLE_PERSON_EMAIL = "support@youguide.com"

# Safety attestation. "Yes" (True) means the product needs NO warnings/safety
# info. The YouGuide activation card DOES carry warnings (not-a-toy, keep from
# under-3s, don't fold the QR), so this is False and we supply the doc instead.
SAFETY_ATTESTATION = False

# GPSR marketplaces (EU only — UK/GB is outside EU GPSR and has its own regime).
# Every SKU is patched in each of these unless you narrow it with --market.
GPSR_MARKETPLACES = ["IE", "FR", "DE", "IT", "ES", "NL", "BE", "SE", "PL"]

# Safety-information document per language, as a DIRECT-DOWNLOAD URL.
# NOTE: compliance_media.source_location must be a direct download link, not a
# "view in browser" page. These are the Google-Drive share links converted to
# their uc?export=download form. Drive serves small PDFs (~45 KB) directly, but
# if Amazon rejects a Drive URL, re-host the PDFs (S3 / your own site) and swap
# the URLs here — that's the only change needed.
_DRIVE = "https://drive.google.com/uc?export=download&id={}"
SAFETY_DOC_URLS = {
    "en": _DRIVE.format("115kJgU2QwctQtkmhNtczMd0Z9iSZ2VoN"),
    "de": _DRIVE.format("1nq955kzrHJQ0kvYQ3m_SCMAHysGhgLhq"),
    "fr": _DRIVE.format("1g2S1PMT572k78x4zs9e7HrSRS-yyo3re"),
    "it": _DRIVE.format("1oujWzK5xndnR0h0uHJUlu_e3vNGXlRn_"),
    "es": _DRIVE.format("1RclPcEaZvDN86JDaReiHD75ihxRfWx30"),
    "nl": _DRIVE.format("14GCI2gQW1trb0LolMu8OF2Z0i4sDigjA"),
    "sv": _DRIVE.format("1PS8lWXa__YE4Sp5-zohNBSotfTT-lhV_"),
    "pl": _DRIVE.format("1_ooAZl8PwAjJhx5PEKDIdJk2Eg3qPCgt"),
}

# Per marketplace: a LIST of (content_language tag, doc key) to attach. Amazon
# shows the document in the shopper's local language, so each store gets the
# matching translation. A marketplace may carry more than one document when it
# is multilingual (Belgium → Dutch primary + French second). compliance_media
# allows one entry per (marketplace, content_type, content_language), so the two
# BE docs differ only by content_language. All tags are validated against the
# live SIM_CARD schema's content_language enum.
MARKET_DOC = {
    "IE": [("en_IE", "en")],
    "FR": [("fr_FR", "fr")],
    "DE": [("de_DE", "de")],
    "IT": [("it_IT", "it")],
    "ES": [("es_ES", "es")],
    "NL": [("nl_NL", "nl")],
    "BE": [("nl_BE", "nl"), ("fr_BE", "fr")],   # Belgium — Dutch primary, French second
    "SE": [("sv_SE", "sv")],
    "PL": [("pl_PL", "pl")],
}

# Throttle — patchListingsItem is 5 req/s. Stay well under it.
SLEEP_BETWEEN_PATCHES = 0.4


# =============================================================================
#  PAYLOAD
# =============================================================================

def gpsr_attributes(marketplace_id: str, docs: list) -> dict:
    """Build the GPSR attribute set for one marketplace.

    `docs` is a list of (content_language, source_url) pairs — one safety
    document per language. Multilingual stores (e.g. Belgium) pass more than one.
    """
    attrs = {
        "gpsr_manufacturer_reference": [
            {"gpsr_manufacturer_email_address": MANUFACTURER_EMAIL,
             "marketplace_id": marketplace_id}
        ],
        "dsa_responsible_party_address": [
            {"value": RESPONSIBLE_PERSON_EMAIL, "marketplace_id": marketplace_id}
        ],
        "gpsr_safety_attestation": [
            {"value": SAFETY_ATTESTATION, "marketplace_id": marketplace_id}
        ],
    }
    media = [
        {"content_type":     "safety_information",
         "content_language": content_language,
         "source_location":  url,
         "marketplace_id":   marketplace_id}
        for content_language, url in docs if url
    ]
    if media:
        attrs["compliance_media"] = media
    return attrs


def build_patches(attrs: dict) -> list:
    """Turn an attribute dict into JSON-PATCH ops. Only top-level attributes can
    be patched, so each op replaces a whole `/attributes/<name>` array."""
    return [{"op": "replace", "path": f"/attributes/{name}", "value": value}
            for name, value in attrs.items()]


# =============================================================================
#  PATCH
# =============================================================================

def patch_gpsr(client, seller_id, sku, marketplace_id, patches, live) -> tuple[bool, str]:
    """Send (or, in dry run, preview) the GPSR patch for one SKU × marketplace."""
    if not live:
        log.info("[DRY]  would PATCH %-22s @ %s  (%d attrs)", sku, marketplace_id, len(patches))
        return True, "DRY_RUN"
    try:
        resp = client.patch_listings_item(
            sellerId=seller_id,
            sku=sku,
            marketplaceIds=[marketplace_id],
            body={"productType": PRODUCT_TYPE, "patches": patches},
            issueLocale="en_US",
        )
        data   = resp.payload or {}
        status = data.get("status", "UNKNOWN")
        issues = data.get("issues", [])
        errors = [i for i in issues if i.get("severity") == "ERROR"]

        if status == "ACCEPTED" and not errors:
            warns = "; ".join(i.get("message", "") for i in issues)
            log.info("[OK]   %-22s @ %s — ACCEPTED%s",
                     sku, marketplace_id, f"  (warnings: {warns})" if warns else "")
            return True, "ACCEPTED"

        summary = "; ".join(i.get("message", str(i)) for i in issues[:10]) or status
        log.warning("[FAIL] %-22s @ %s — %s | %s", sku, marketplace_id, status, summary)
        return False, summary

    except SellingApiException as exc:
        # A 404 here means the SKU isn't live in this marketplace — expected for
        # some sku/market combos; log it as a skip rather than an error.
        code = getattr(exc, "code", None)
        if code == 404:
            log.info("[SKIP] %-22s @ %s — not listed here (404)", sku, marketplace_id)
            return False, "NOT_LISTED"
        log.error("[ERR]  %-22s @ %s — %s", sku, marketplace_id, exc)
        return False, f"ERR {exc}"


# =============================================================================
#  SHEET
# =============================================================================

def read_skus(only_sku: str = "") -> list:
    """Read the unique child SKUs (column A) from the sheet, in order."""
    if only_sku:
        return [only_sku]
    ws   = get_worksheet()
    rows = ws.get_all_values()[1:]      # drop header row
    skus, seen = [], set()
    for row in rows:
        sku = (row[COL_SKU] if len(row) > COL_SKU else "").strip()
        if sku and sku not in seen:
            seen.add(sku)
            skus.append(sku)
    return skus


# =============================================================================
#  MAIN
# =============================================================================

def main() -> None:
    ap = argparse.ArgumentParser(description="Bulk-apply GPSR compliance to Amazon listings.")
    ap.add_argument("--live",   action="store_true", help="actually send the patches (default: dry run)")
    ap.add_argument("--sku",    default="", help="patch only this one SKU")
    ap.add_argument("--market", default="", help="patch only this marketplace code (e.g. DE)")
    ap.add_argument("--limit",  type=int, default=0, help="patch only the first N SKUs")
    args = ap.parse_args()

    # Which marketplaces?
    if args.market:
        code = args.market.strip().upper()
        if code not in GPSR_MARKETPLACES:
            log.error("Marketplace %s is not in GPSR_MARKETPLACES %s", code, GPSR_MARKETPLACES)
            sys.exit(1)
        markets = [code]
    else:
        markets = GPSR_MARKETPLACES

    # Which SKUs?
    skus = read_skus(args.sku)
    if args.limit:
        skus = skus[: args.limit]
    if not skus:
        log.error("No SKUs found in the sheet.")
        sys.exit(1)

    mode = "*** LIVE — patching Amazon ***" if args.live else "DRY RUN (no writes; --live to apply)"
    log.info("GPSR sync  %s", mode)
    log.info("SKUs: %d   Markets: %s   → %d patch calls",
             len(skus), ",".join(markets), len(skus) * len(markets))

    # Pre-build one client per marketplace (skip markets whose token is unset).
    clients = {}
    for code in markets:
        client = get_client_for_marketplace(code)
        if client is None:
            log.warning("No SP-API token for %s (region %s) — skipping that marketplace.",
                        code, MARKETPLACES[code]["region"])
            continue
        clients[code] = client
    if not clients:
        log.error("No usable marketplace clients — check your .env refresh tokens.")
        sys.exit(1)

    # Tally results per marketplace.
    tally = {code: {"ok": 0, "fail": 0, "skip": 0} for code in clients}

    for code, client in clients.items():
        seller_id = SELLER_IDS.get(MARKETPLACES[code]["token"], "")
        docs = []
        for content_language, doc_key in MARKET_DOC[code]:
            url = SAFETY_DOC_URLS.get(doc_key, "")
            if not url:
                log.warning("No safety document URL for %s (lang %s) — skipping that document.",
                            code, doc_key)
                continue
            docs.append((content_language, url))
        mp_id   = MARKETPLACES[code]["id"]
        patches = build_patches(gpsr_attributes(mp_id, docs))

        log.info("── %s (%s) ──────────────────────────────", code, mp_id)
        for sku in skus:
            ok, status = patch_gpsr(client, seller_id, sku, mp_id, patches, args.live)
            if ok:
                tally[code]["ok"] += 1
            elif status == "NOT_LISTED":
                tally[code]["skip"] += 1
            else:
                tally[code]["fail"] += 1
            if args.live:
                time.sleep(SLEEP_BETWEEN_PATCHES)

    log.info("================ SUMMARY ================")
    for code, t in tally.items():
        log.info("%s:  ok=%d  fail=%d  skip=%d", code, t["ok"], t["fail"], t["skip"])
    if not args.live:
        log.info("Dry run only — re-run with --live to apply.")


if __name__ == "__main__":
    main()
