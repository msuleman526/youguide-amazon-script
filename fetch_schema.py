#!/usr/bin/env python3
"""
Fetch the authoritative SIM_CARD product-type JSON schema from Amazon so we
know the exact required attribute keys, enums, and formats for the UK
marketplace. Saves the full schema and prints the required attributes.
"""

import json
import urllib.request

from sp_api.api import ProductTypeDefinitions
from upload_listings import SP_API_CREDENTIALS, MARKETPLACE, MARKETPLACE_ID, PRODUCT_TYPE


def fetch(requirements: str, out_file: str) -> dict:
    print(f"\n=== {PRODUCT_TYPE} schema (requirements={requirements}) ===")
    res = ProductTypeDefinitions(credentials=SP_API_CREDENTIALS, marketplace=MARKETPLACE).get_definitions_product_type(
        productType=PRODUCT_TYPE,
        marketplaceIds=[MARKETPLACE_ID],
        requirements=requirements,
        locale="en_GB",
    )
    meta = res.payload
    link = meta["schema"]["link"]["resource"]
    with urllib.request.urlopen(link) as r:
        schema = json.loads(r.read().decode("utf-8"))

    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(schema, f, indent=2)
    print(f"Saved full schema -> {out_file}")

    required = schema.get("required", [])
    props = schema.get("properties", {})
    print(f"\nTop-level REQUIRED attributes ({len(required)}):\n")
    for key in sorted(required):
        p = props.get(key, {})
        title = p.get("title", "")
        # Try to surface enum values / sub-structure hints
        hint = ""
        items = p.get("items", {})
        sub_props = items.get("properties", {})
        if "value" in sub_props and "enum" in sub_props["value"]:
            hint = " enum=" + str(sub_props["value"]["enum"][:8])
        elif sub_props:
            hint = " fields=" + str(list(sub_props.keys()))
        print(f"  {key:45} | {title}{hint}")
    return schema


if __name__ == "__main__":
    fetch("LISTING", "sim_card_schema_listing.json")
    fetch("LISTING_PRODUCT_ONLY", "sim_card_schema_product_only.json")
    print("\nDone. Open the JSON files for full detail on any attribute.")
