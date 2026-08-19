#!/usr/bin/env python3
"""
gpsr_status.py — read-only status check for the GPSR rollout
============================================================
Queries Amazon LIVE (get_listings_item) and reports, per SKU × marketplace:
  * GPSR   — are all four GPSR attributes present? (or which are MISSING)
  * STATUS — the listing's live status (BUYABLE / DISCOVERABLE / ...)
  * ISSUES — count of ERROR-severity issues, with the first messages

It never writes anything — safe to run at any time, including while
upload_gpsr.py is still patching in another tab. Use it to watch progress and
confirm listings went compliant / active.

Usage:
    python gpsr_status.py --sku HAJJ-LUX-10GB-30D --market BE   # one sku, one market
    python gpsr_status.py --market BE                           # whole BE marketplace
    python gpsr_status.py --market DE --limit 10                # first 10 SKUs on DE
    python gpsr_status.py                                       # every SKU × every GPSR market (slow)
"""

import argparse
import sys
import time

from sp_api.base import SellingApiException

from upload_listings import (
    log,
    MARKETPLACES,
    SELLER_IDS,
    get_client_for_marketplace,
)
from upload_gpsr import GPSR_MARKETPLACES, read_skus

# The four attributes upload_gpsr.py writes. All four present → fully compliant.
GPSR_ATTRS = [
    "gpsr_manufacturer_reference",
    "dsa_responsible_party_address",
    "gpsr_safety_attestation",
    "compliance_media",
]

SLEEP_BETWEEN_READS = 0.3   # get_listings_item is 5 req/s; stay well under it.


def check_one(client, seller_id, sku, mp_id) -> dict:
    """Read one SKU in one marketplace. Returns a status dict (never raises)."""
    try:
        resp = client.get_listings_item(
            sellerId=seller_id,
            sku=sku,
            marketplaceIds=[mp_id],
            includedData=["summaries", "attributes", "issues"],
        )
    except SellingApiException as exc:
        if getattr(exc, "code", None) == 404:
            return {"listed": False}
        return {"error": str(exc)}

    payload   = resp.payload or {}
    attrs     = payload.get("attributes") or {}
    summaries = payload.get("summaries") or []
    issues    = payload.get("issues") or []

    missing = [a for a in GPSR_ATTRS if a not in attrs]

    status = ""
    if summaries:
        st = summaries[0].get("status")
        status = ",".join(st) if isinstance(st, list) else (st or "")

    errors = [i for i in issues if i.get("severity") == "ERROR"]

    return {
        "listed":  True,
        "missing": missing,
        "status":  status or "UNKNOWN",
        "errors":  errors,
    }


def main() -> None:
    ap = argparse.ArgumentParser(description="Read-only GPSR/status check for Amazon listings.")
    ap.add_argument("--sku",    default="", help="check only this one SKU")
    ap.add_argument("--market", default="", help="check only this marketplace code (e.g. BE)")
    ap.add_argument("--limit",  type=int, default=0, help="check only the first N SKUs")
    args = ap.parse_args()

    if args.market:
        code = args.market.strip().upper()
        if code not in GPSR_MARKETPLACES:
            log.error("Marketplace %s not in GPSR_MARKETPLACES %s", code, GPSR_MARKETPLACES)
            sys.exit(1)
        markets = [code]
    else:
        markets = GPSR_MARKETPLACES

    skus = read_skus(args.sku)
    if args.limit:
        skus = skus[: args.limit]
    if not skus:
        log.error("No SKUs found in the sheet.")
        sys.exit(1)

    log.info("GPSR status check — SKUs: %d   Markets: %s   → %d reads",
             len(skus), ",".join(markets), len(skus) * len(markets))

    tally = {code: {"compliant": 0, "partial": 0, "missing": 0,
                    "buyable": 0, "errors": 0, "not_listed": 0} for code in markets}

    for code in markets:
        client = get_client_for_marketplace(code)
        if client is None:
            log.warning("No SP-API token for %s — skipping.", code)
            continue
        seller_id = SELLER_IDS.get(MARKETPLACES[code]["token"], "")
        mp_id     = MARKETPLACES[code]["id"]
        log.info("== %s (%s) ==", code, mp_id)

        for sku in skus:
            r = check_one(client, seller_id, sku, mp_id)
            time.sleep(SLEEP_BETWEEN_READS)

            if r.get("error"):
                log.error("[%s] %-22s  READ ERROR: %s", code, sku, r["error"])
                continue
            if not r["listed"]:
                tally[code]["not_listed"] += 1
                log.info("[%s] %-22s  not listed here (404)", code, sku)
                continue

            missing = r["missing"]
            if not missing:
                gpsr = "GPSR:full"
                tally[code]["compliant"] += 1
            elif len(missing) == len(GPSR_ATTRS):
                gpsr = "GPSR:NONE"
                tally[code]["missing"] += 1
            else:
                gpsr = "GPSR:PARTIAL missing=" + ",".join(missing)
                tally[code]["partial"] += 1

            status = r["status"]
            if "BUYABLE" in status:
                tally[code]["buyable"] += 1

            errs = r["errors"]
            if errs:
                tally[code]["errors"] += 1
            err_txt = ""
            if errs:
                err_txt = "  ERRORS(%d): %s" % (
                    len(errs), "; ".join(e.get("message", "") for e in errs[:3]))

            log.info("[%s] %-22s  %-45s  status=%-22s%s",
                     code, sku, gpsr, status, err_txt)

    log.info("================ SUMMARY ================")
    for code, t in tally.items():
        log.info("%s: compliant=%d partial=%d none=%d | buyable=%d with-errors=%d not-listed=%d",
                 code, t["compliant"], t["partial"], t["missing"],
                 t["buyable"], t["errors"], t["not_listed"])


if __name__ == "__main__":
    main()
