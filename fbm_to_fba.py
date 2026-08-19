#!/usr/bin/env python3
"""
Convert YouGuide listings from FBM (Merchant Fulfilled) to FBA (Amazon Fulfilled)
via the SP-API Listings Items API.

WHAT IT DOES
    For each SKU it PATCHes `fulfillment_availability`: adds the marketplace's
    Amazon fulfilment channel ("AMAZON_EU" on UK/EU) and deletes the "DEFAULT"
    (merchant) entry. Quantity is deliberately NOT sent: under FBA the quantity
    is owned by Amazon and derives from units physically received at a
    fulfilment centre.

WHAT IT DOES NOT DO
    Flipping the channel converts the OFFER only. It does not send stock. The
    listing shows 0 / out-of-stock until an FBA inbound shipment is received.
    Amazon's own schema note: "Specifying a value other than DEFAULT will cancel
    the Merchant-fulfilled offering."

THREE AMAZON BEHAVIOURS THIS SCRIPT WORKS AROUND (all verified live, 2026-07-17,
on AALAND-ISLANDS-3GB-15D — do not "simplify" these away):

  1. fulfillment_availability is SELECTOR-KEYED on fulfillment_channel_code.
     A "replace" op does NOT replace the array — it merges, leaving the SKU on
     BOTH channels and still fulfilled by DEFAULT. The merchant entry must be
     deleted explicitly, by value. A delete op with no "value" is rejected:
     "Invalid empty value provided in patch at index of 0".

  2. The DEFAULT entry LINGERS in `attributes` even after a successful delete.
     The authoritative live state is the top-level `fulfillmentAvailability`
     (offer) view, not the attributes list. Verify against the offer.

  3. The offer view takes ~30-60s to reflect a patch. A read-back at +20s still
     reported DEFAULT on a SKU that had in fact converted. Hence the two-phase
     design: patch everything, wait once, then verify everything. Verifying
     per-SKU with a 60s wait would turn a 30-minute job into a 7-hour one.

Also note: `mode=VALIDATION_PREVIEW` validates a patch document in ISOLATION,
without merging it against the live listing, so a channel-only patch is rejected
with "90220 'Are batteries required?' is required but missing" even though the
listing carries that attribute. ECHO_ATTRS echoes it back unchanged to satisfy
the preview validator; it is a no-op against the live listing.

USAGE
    python fbm_to_fba.py --sku AALAND-ISLANDS-3GB-15D        # preview one SKU
    python fbm_to_fba.py --limit 10                          # preview first 10 from Excel
    python fbm_to_fba.py                                     # preview ALL from Excel
    python fbm_to_fba.py --sku AALAND-ISLANDS-3GB-15D --live # APPLY one SKU
    python fbm_to_fba.py --live                              # APPLY all
    python fbm_to_fba.py --markets UK,DE --live              # multiple marketplaces
    python fbm_to_fba.py --verify-only                       # just audit current channels

Every run writes a timestamped .log and .csv report into logs/ for the client.
"""

import argparse
import csv
import datetime as _dt
import logging
import os
import sys
import time
from typing import Optional

from sp_api.api import ListingsItems
from sp_api.base import SellingApiException

# Windows consoles default to cp1252 and cannot render the log's box/arrow glyphs.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from upload_listings import (
    FBM_CHANNEL,
    LWA_CREDENTIALS,
    MARKETPLACES,
    PRODUCT_TYPE,
    REFRESH_TOKENS,
    REGION_ENDPOINT,
    SELLER_IDS,
    fba_channel_for,
)

# ── Configuration ────────────────────────────────────────────────────────────

EXCEL_PATH = os.path.join("FBM to FBA",
                          "YOUGUIDE poducts FOR every MARKET place 1ST DISTRIBUTION (1).xlsx")
EXCEL_SHEET = "Youguide eSIM"
EXCEL_SKU_COL = 1        # column B — SKU
EXCEL_GTIN_COL = 0       # column A — GTIN (reported only, not sent)

# The target fulfilment network per marketplace comes from upload_listings
# (fba_channel_for / FBM_CHANNEL) so the converter and the sync loop can never
# disagree about what "FBA" means for a given marketplace.

# Echoed back unchanged so VALIDATION_PREVIEW does not report them missing.
ECHO_ATTRS = ["batteries_required"]

SLEEP_BETWEEN_CALLS = 0.6   # SP-API patch is 5 rps; stay well under it.
API_RETRIES = 3             # transient ReadTimeouts are common under throttling
RETRY_BACKOFF = 8.0
HTTP_TIMEOUT = 60.0         # Amazon's patch validation can exceed the library default
# Wait before the verify pass. 90 s was measured once and is TOO SHORT: on
# 2026-08-18 a patch took ~7 minutes to reach the offer view, so a 90 s verify
# reported NOT_APPLIED for hundreds of SKUs that were merely still in flight —
# indistinguishable from the genuine silent no-op being hunted at the time.
# Err long: a false NOT_APPLIED costs a needless re-run of everything.
SETTLE_SECONDS = 420.0

LOG_DIR = "logs"

# Statuses that mean "this SKU is now FBA".
OK_STATUSES = ("CONVERTED", "ALREADY_FBA")


def make_client(market: str) -> Optional[ListingsItems]:
    """ListingsItems client for a marketplace's auth group, or None if no token.

    Mirrors upload_listings.get_client_for_marketplace() but sets an explicit
    HTTP timeout, which that helper does not expose.
    """
    cfg = MARKETPLACES[market]
    token = REFRESH_TOKENS.get(cfg["token"], "")
    if not token:
        return None
    return ListingsItems(
        credentials={"refresh_token": token, **LWA_CREDENTIALS},
        marketplace=REGION_ENDPOINT[cfg["region"]],
        timeout=HTTP_TIMEOUT,
    )


def with_retry(fn, *args, **kwargs):
    """Call an SP-API method, retrying transient network/throttle timeouts."""
    last = None
    for attempt in range(API_RETRIES):
        try:
            return fn(*args, **kwargs)
        except SellingApiException:
            raise                       # a real API verdict — let the caller handle it
        except Exception as exc:        # ReadTimeout, connection resets, etc.
            last = exc
            if attempt < API_RETRIES - 1:
                time.sleep(RETRY_BACKOFF * (attempt + 1))
    raise last


def build_logger(tag: str) -> tuple[logging.Logger, str, str]:
    """Console + file logger. Returns (logger, log_path, csv_path)."""
    os.makedirs(LOG_DIR, exist_ok=True)
    stamp = _dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    log_path = os.path.join(LOG_DIR, f"fbm_to_fba_{tag}_{stamp}.log")
    csv_path = os.path.join(LOG_DIR, f"fbm_to_fba_{tag}_{stamp}.csv")

    logger = logging.getLogger("fbm_to_fba")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    logger.propagate = False
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s",
                            datefmt="%Y-%m-%d %H:%M:%S")
    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    logger.addHandler(sh)

    logging.getLogger("httpx").setLevel(logging.WARNING)   # one INFO line per call
    return logger, log_path, csv_path


def read_skus_from_excel(path: str, limit: Optional[int]) -> list:
    """Read (sku, gtin, desc) rows from the distribution workbook."""
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[EXCEL_SHEET]
    rows = []
    for raw in ws.iter_rows(min_row=2, values_only=True):
        if not raw or len(raw) <= EXCEL_SKU_COL:
            continue
        sku = raw[EXCEL_SKU_COL]
        if not sku or not str(sku).strip():
            continue                      # skips the trailing =SUM() row
        rows.append({
            "sku":  str(sku).strip(),
            "gtin": str(raw[EXCEL_GTIN_COL]).strip() if raw[EXCEL_GTIN_COL] else "",
            "desc": str(raw[2]).strip() if len(raw) > 2 and raw[2] else "",
        })
        if limit and len(rows) >= limit:
            break
    wb.close()
    return rows


def attr_channels(payload: dict) -> list:
    """Fulfilment channel codes present in the listing's ATTRIBUTES.

    Informational only. A converted SKU keeps a residual DEFAULT entry here
    (see behaviour 2 in the module docstring), so never treat this as the
    live state — use offer_channel().
    """
    fa = payload.get("attributes", {}).get("fulfillment_availability") or []
    return [e.get("fulfillment_channel_code", "") for e in fa]


def offer_channel(payload: dict) -> tuple[str, Optional[int]]:
    """The channel Amazon reports on the OFFER — the authoritative live state."""
    fa = payload.get("fulfillmentAvailability") or []
    if not fa:
        return "", None
    return fa[0].get("fulfillmentChannelCode", ""), fa[0].get("quantity")


def format_issues(issues: list) -> str:
    """Flatten Amazon's issue list into one readable line."""
    parts = []
    for i in issues or []:
        msg = (i.get("message", "") or "").replace("\n", " ").strip()
        parts.append(f"[{i.get('severity','?')} {i.get('code','?')}] {msg}")
    return " ; ".join(parts)


def read_listing(client, seller_id, marketplace_id, sku):
    return with_retry(
        client.get_listings_item,
        sellerId=seller_id, sku=sku, marketplaceIds=[marketplace_id],
        includedData=["summaries", "attributes", "fulfillmentAvailability", "issues"],
    )


def submit_one(client, log, seller_id, marketplace_id, market, sku, live) -> dict:
    """Read a SKU and submit its channel patch. Does NOT verify (see main)."""
    result = {"sku": sku, "market": market, "asin": "", "before": "",
              "after": "", "status": "", "issues": ""}
    channel = fba_channel_for(market)

    # ── Read current state ───────────────────────────────────────────────────
    try:
        resp = read_listing(client, seller_id, marketplace_id, sku)
    except SellingApiException as exc:
        if getattr(exc, "code", None) == 404:
            result.update(status="NOT_LISTED", issues="SKU not found in this marketplace")
            log.warning("  %-28s %s  NOT_LISTED - no such SKU here", sku, market)
        else:
            result.update(status="READ_ERROR", issues=str(exc))
            log.error("  %-28s %s  READ_ERROR: %s", sku, market, exc)
        return result
    except Exception as exc:
        result.update(status="READ_TIMEOUT", issues=f"{type(exc).__name__}: {exc}")
        log.error("  %-28s %s  READ_TIMEOUT after %d retries - skipped, rerun later",
                  sku, market, API_RETRIES)
        return result

    payload = resp.payload or {}
    result["asin"] = (payload.get("summaries") or [{}])[0].get("asin", "")
    codes = attr_channels(payload)
    live_channel, _qty = offer_channel(payload)
    result["before"] = live_channel or ",".join(codes) or "(none)"

    if live_channel == channel:
        result.update(status="ALREADY_FBA", after=channel)
        log.info("  %-28s %s  ALREADY_FBA - no change needed", sku, market)
        return result

    unexpected = [c for c in codes if c not in (channel, FBM_CHANNEL)]
    if unexpected:
        result.update(status="UNEXPECTED_CHANNEL", after=",".join(codes))
        log.warning("  %-28s %s  UNEXPECTED_CHANNEL %s - skipped",
                    sku, market, ",".join(unexpected))
        return result

    # ── Build the patch ──────────────────────────────────────────────────────
    # Add the FBA entry, then delete the merchant one. Order matters: the schema
    # requires at least one entry, so never delete before adding.
    #
    # TWO BUGS LIVED HERE UNTIL 2026-08-18. Both made Amazon answer ACCEPTED and
    # then change nothing, which is the worst possible failure: a run of 624 SKUs
    # reported success and applied 0. Do not "simplify" either fix away.
    #
    #   1. The add was skipped when `channel` already appeared in `codes`. But a
    #      residual FBA entry in `attributes` does NOT mean the OFFER is on that
    #      channel (behaviour 2 in the module docstring), so the patch collapsed
    #      to a lone delete. Always re-assert the FBA entry; re-adding one that is
    #      already there is harmless.
    #   2. The delete synthesized its own value, `{"fulfillment_channel_code":
    #      "DEFAULT"}`, which matches NOTHING when the stored entry also carries a
    #      quantity — as it does for anything market_activation.py switched off,
    #      where the entry is `{"fulfillment_channel_code": "DEFAULT",
    #      "quantity": 0}`. Delete BY THE EXACT STORED VALUE, read back from the
    #      live listing. Verified live on AUSTRALIA-3GB-15D: offer went DEFAULT ->
    #      AMAZON_EU and status DISCOVERABLE -> DISCOVERABLE,BUYABLE.
    fa_entries = payload.get("attributes", {}).get("fulfillment_availability") or []
    patches = [{"op": "add", "path": "/attributes/fulfillment_availability",
                "value": [{"fulfillment_channel_code": channel}]}]
    merchant = [e for e in fa_entries
                if e.get("fulfillment_channel_code") == FBM_CHANNEL]
    if merchant:
        patches.append({"op": "delete", "path": "/attributes/fulfillment_availability",
                        "value": merchant})

    attrs = payload.get("attributes", {})
    for name in ECHO_ATTRS:
        if attrs.get(name):
            patches.append({"op": "replace", "path": f"/attributes/{name}",
                            "value": attrs[name]})

    params = {
        "sellerId": seller_id, "sku": sku, "marketplaceIds": [marketplace_id],
        "includedData": ["issues"],
        "body": {"productType": PRODUCT_TYPE, "patches": patches},
    }
    if not live:
        params["mode"] = "VALIDATION_PREVIEW"

    try:
        resp = with_retry(client.patch_listings_item, **params)
    except SellingApiException as exc:
        result.update(status="API_ERROR", issues=str(exc))
        log.error("  %-28s %s  API_ERROR: %s", sku, market, exc)
        return result
    except Exception as exc:
        result.update(status="PATCH_TIMEOUT", issues=f"{type(exc).__name__}: {exc}")
        log.error("  %-28s %s  PATCH_TIMEOUT - state unconfirmed, verify pass will "
                  "report the truth", sku, market)
        return result

    out = resp.payload or {}
    amz_status = (out.get("status") or "").upper()
    result["issues"] = format_issues(out.get("issues"))
    errors = [i for i in (out.get("issues") or []) if i.get("severity") == "ERROR"]

    if amz_status in ("ACCEPTED", "VALID") and not errors:
        # ACCEPTED means queued, not applied — the verify pass decides.
        result.update(status="PREVIEW_OK" if not live else "SUBMITTED", after=channel)
        log.info("  %-28s %s  %-12s %s -> %s%s", sku, market, result["status"],
                 result["before"], channel,
                 f"  | {result['issues']}" if result["issues"] else "")
    else:
        result.update(status="REJECTED", after=result["before"])
        log.error("  %-28s %s  REJECTED (%s): %s", sku, market, amz_status or "?",
                  result["issues"] or "no detail returned")

    time.sleep(SLEEP_BETWEEN_CALLS)
    return result


def verify_one(client, log, seller_id, marketplace_id, market, rec) -> None:
    """Re-read a SKU and record whether the offer really moved to FBA."""
    channel = fba_channel_for(market)
    sku = rec["sku"]
    try:
        payload = read_listing(client, seller_id, marketplace_id, sku).payload or {}
    except Exception as exc:
        rec["status"] = "UNVERIFIED"
        rec["issues"] = f"submitted, but read-back failed: {type(exc).__name__}"
        log.warning("  %-28s %s  UNVERIFIED - read-back failed", sku, market)
        return

    live_channel, qty = offer_channel(payload)
    rec["after"] = live_channel or "(none)"

    if live_channel == channel:
        rec["status"] = "CONVERTED"
        log.info("  %-28s %s  CONVERTED    offer=%s qty=%s", sku, market,
                 live_channel, qty if qty is not None else "-")
    else:
        rec["status"] = "NOT_APPLIED"
        rec["issues"] = (f"accepted but offer still reports {live_channel or '(none)'}; "
                         f"attributes={attr_channels(payload)}")
        log.error("  %-28s %s  NOT_APPLIED  offer still %s", sku, market,
                  live_channel or "(none)")
    time.sleep(SLEEP_BETWEEN_CALLS)


def main() -> None:
    ap = argparse.ArgumentParser(description="Convert Amazon listings FBM -> FBA via SP-API")
    ap.add_argument("--sku", action="append", help="only this SKU (repeatable)")
    ap.add_argument("--markets", default="UK", help="comma list of marketplace codes (default UK)")
    ap.add_argument("--limit", type=int, help="only the first N SKUs from the Excel")
    ap.add_argument("--excel", default=EXCEL_PATH, help="path to the distribution workbook")
    ap.add_argument("--live", action="store_true",
                    help="APPLY changes. Without this flag Amazon only validates (no change).")
    ap.add_argument("--verify-only", action="store_true",
                    help="audit the current fulfilment channel; change nothing")
    args = ap.parse_args()

    tag = "verify" if args.verify_only else ("live" if args.live else "preview")
    log, log_path, csv_path = build_logger(tag)

    markets = [m.strip().upper() for m in args.markets.split(",") if m.strip()]
    bad = [m for m in markets if m not in MARKETPLACES]
    if bad:
        log.error("Unknown marketplace code(s): %s", ", ".join(bad))
        raise SystemExit(2)

    if args.sku:
        products = [{"sku": s, "gtin": "", "desc": ""} for s in args.sku]
        source = "--sku argument"
    else:
        products = read_skus_from_excel(args.excel, args.limit)
        source = args.excel

    mode = ("AUDIT ONLY - nothing will be changed" if args.verify_only else
            "LIVE - CHANGES WILL BE APPLIED" if args.live else
            "VALIDATION PREVIEW - nothing will be changed")

    log.info("=" * 78)
    log.info("YouGuide - Amazon FBM -> FBA conversion")
    log.info("=" * 78)
    log.info("Mode         : %s", mode)
    log.info("Source       : %s", source)
    log.info("SKUs         : %d", len(products))
    log.info("Marketplaces : %s", ", ".join(markets))
    log.info("Change       : fulfillment_availability.fulfillment_channel_code %s -> %s",
             FBM_CHANNEL, "/".join(sorted({fba_channel_for(m) for m in markets})))
    log.info("Note         : quantity is not sent - under FBA Amazon owns the quantity.")
    log.info("               Listings stay out-of-stock until inbound units are received.")
    log.info("=" * 78)

    results = []
    for market in markets:
        client = make_client(market)
        if client is None:
            log.error("%s: no refresh token configured for its region - skipped", market)
            continue
        cfg = MARKETPLACES[market]
        seller_id = SELLER_IDS.get(cfg["token"], "")
        channel = fba_channel_for(market)

        log.info("")
        log.info("-- Marketplace %s (%s) --", market, cfg["id"])

        if args.verify_only:
            for n, prod in enumerate(products, 1):
                rec = {"sku": prod["sku"], "market": market, "asin": "",
                       "before": "", "after": "", "status": "", "issues": ""}
                log.info("[%d/%d]", n, len(products))
                verify_one(client, log, seller_id, cfg["id"], market, rec)
                results.append(rec)
            continue

        # ── Phase 1: submit ──────────────────────────────────────────────────
        log.info("Phase 1/2 - submitting channel changes")
        market_results = []
        for n, prod in enumerate(products, 1):
            log.info("[%d/%d]", n, len(products))
            market_results.append(
                submit_one(client, log, seller_id, cfg["id"], market, prod["sku"], args.live)
            )
        results.extend(market_results)

        # ── Phase 2: verify ──────────────────────────────────────────────────
        # Only live runs change anything, and only SUBMITTED rows need checking.
        pending = [r for r in market_results if r["status"] == "SUBMITTED"]
        if args.live and pending:
            log.info("")
            log.info("Phase 2/2 - waiting %ds for Amazon's offer view to settle, then "
                     "verifying %d SKU(s)", int(SETTLE_SECONDS), len(pending))
            time.sleep(SETTLE_SECONDS)
            for n, rec in enumerate(pending, 1):
                log.info("[verify %d/%d]", n, len(pending))
                verify_one(client, log, seller_id, cfg["id"], market, rec)

    # ── Summary ──────────────────────────────────────────────────────────────
    counts: dict = {}
    for r in results:
        counts[r["status"]] = counts.get(r["status"], 0) + 1

    log.info("")
    log.info("=" * 78)
    log.info("SUMMARY - %s", mode)
    log.info("=" * 78)
    for status, n in sorted(counts.items(), key=lambda kv: -kv[1]):
        log.info("  %-20s %d", status, n)
    log.info("  %-20s %d", "TOTAL", len(results))

    if args.live:
        ok = sum(n for s, n in counts.items() if s in OK_STATUSES)
        log.info("")
        log.info("  %d of %d listing(s) are now FBA (%s).", ok, len(results),
                 ", ".join(sorted({fba_channel_for(m) for m in markets})))
        failed = [r for r in results if r["status"] not in OK_STATUSES]
        if failed:
            log.info("  %d need attention:", len(failed))
            for r in failed:
                log.info("    %-28s %s  %s  %s", r["sku"], r["market"],
                         r["status"], r["issues"][:90])

    with open(csv_path, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=["sku", "market", "asin", "before",
                                           "after", "status", "issues"])
        w.writeheader()
        w.writerows(results)

    log.info("")
    log.info("Log report : %s", log_path)
    log.info("CSV report : %s", csv_path)
    if not args.live and not args.verify_only:
        log.info("")
        log.info("This was a PREVIEW - no listing was modified. Re-run with --live to apply.")


if __name__ == "__main__":
    main()
