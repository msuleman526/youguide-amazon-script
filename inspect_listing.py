#!/usr/bin/env python3
"""
Field-extraction test for the reverse sync (Amazon → Sheet).

Reads ONE SKU from Amazon, prints the raw payload Amazon returns, then prints
exactly what `read_amazon_state()` extracts from it. Use this to confirm the
reverse-sync parsing maps price / quantity / title / image correctly for your
real listings BEFORE turning DRY_RUN off in upload_listings.py.

This is read-only — it never creates, updates, or deletes anything.

Usage:
    python inspect_listing.py                 # uses the default SKU below
    python inspect_listing.py ZAMBIA-3GB-15D  # any of your SKUs
"""

import json
import sys

from sp_api.api import ListingsItems
from sp_api.base import SellingApiException

from upload_listings import (
    SP_API_CREDENTIALS,
    MARKETPLACE,
    MARKETPLACE_ID,
    SELLER_ID,
    read_amazon_state,
)

DEFAULT_SKU = "HAJJ-LUX-10GB-30D"


def main() -> None:
    sku = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_SKU
    client = ListingsItems(credentials=SP_API_CREDENTIALS, marketplace=MARKETPLACE)

    print(f"Reading SKU: {sku}  (marketplace {MARKETPLACE.name} / {MARKETPLACE_ID})\n")

    try:
        resp = client.get_listings_item(
            sellerId=SELLER_ID,
            sku=sku,
            marketplaceIds=[MARKETPLACE_ID],
            includedData=["summaries", "attributes", "offers",
                          "fulfillmentAvailability", "relationships", "issues"],
        )
    except SellingApiException as exc:
        print(f"API error: {exc}")
        if getattr(exc, "code", None) == 404:
            print(">> 404 — this SKU is not currently listed on Amazon.")
        raise SystemExit(1)

    print("=== RAW PAYLOAD (first 6000 chars) ===")
    print(json.dumps(resp.payload, indent=2)[:6000])

    print("\n=== PARSED BY read_amazon_state() ===")
    state = read_amazon_state(client, sku)
    print(json.dumps(state, indent=2))

    prod = state.get("product")
    if prod:
        print("\n=== KEY FIELDS the reverse sync would write back ===")
        for key in ("title", "brand", "variation", "price", "quantity", "image1", "parent_sku"):
            print(f"  {key:12} = {prod.get(key)!r}")
    else:
        print("\n(No product extracted — SKU not found or empty payload.)")


if __name__ == "__main__":
    main()
