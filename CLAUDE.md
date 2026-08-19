# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A set of standalone Python scripts that keep a Google Sheet and Amazon listings **in two-way sync** for YouGuide travel eSIMs / SIM cards, via Amazon's Selling Partner API (SP-API). One sheet row = one product; its **countries** column names every marketplace to sell it in, so a single row fans out to many Amazon marketplaces (UK/IE/FR/DE/IT/ES/NL/BE/SE/PL/US/CA/MX/JP/AU) under the same SKU — no extra rows. There is no package, build step, or test suite — each script is run directly.

## Commands

```powershell
pip install python-amazon-sp-api gspread google-auth   # one-time dependency install

python upload_listings.py    # core: two-way sync loop (Sheet ⇄ Amazon), runs forever, polls every 5–10 min
python fetch_listings.py     # read-only: pull EVERY live listing via Reports API into the "amazon_current" tab
python fetch_schema.py       # refresh sim_card_schema_*.json from Amazon's live SIM_CARD product-type schema
python inspect_listing.py [SKU]   # read-only: dump one SKU's raw Amazon payload + what read_amazon_state() extracts
python fbm_to_fba.py              # convert existing listings FBM -> FBA (previews by default; --live applies)
python market_activation.py       # keep ONLY UK selling: deactivate (never delete) every other marketplace + write the audit .xlsx (previews by default; --live applies)
python upload_listings.py --rates [COST]    # read-only: tier markup + today's FX = what a USD cost sells for per marketplace
python fetch_esim_prices.py                 # preview: match every row to an eSIM Access package BY NAME, show its USD cost
python fetch_esim_prices.py --apply         # write column G (cost) + X (packageCode) from the supplier catalogue
python fetch_esim_prices.py --find <name>   # search the supplier catalogue (to pin a row's packageCode by hand)
python fetch_orders.py            # export every SE + PL order to Sheets/orders_<CODE>_<stamp>.xlsx (read-only)
python fetch_orders.py --markets SE --since 2026-01-01   # one marketplace / narrower window
python diagnose.py           # auth troubleshooter — verifies token + Product Listing role without changing listings
```

There are no lint/test commands. To syntax-check before running:
`python -c "import ast; ast.parse(open('upload_listings.py', encoding='utf-8').read())"`

**First run** of any script opens a browser for Google OAuth and caches the token to `authorized_user.json`. `upload_listings.py` runs an infinite polling loop — stop it with Ctrl-C.

**`DRY_RUN`** (top of `upload_listings.py`, currently `False`): when `True`, no writes go to Amazon or the sheet — it only logs the create/update/delete/pull decisions it *would* make (reads still happen). Flip to `True` to safely rehearse changes.

**Secrets** live in `.env` (gitignored), loaded by a tiny built-in `load_dotenv()` — python-dotenv is **not** a dependency. Copy `.env.example` → `.env` and fill in values. `authorized_user.json`, `oauth_client.json` hold Google OAuth tokens/secrets. Treat all as sensitive.

## Architecture

`upload_listings.py` is the core; the other scripts import its config/constants/functions so **credentials and IDs live in exactly one place** (the credentials block near the top, sourced from `.env`). Change config in `.env`, not in the dependent scripts.

**Config resolution:** `.env` → `os.environ` → module constants. LWA app id/secret are shared across regions (one app); the **refresh token AND seller ID differ per SP-API authorization group** (`EU`, `NA`, `JP`, `AU`). `REFRESH_TOKENS` / `SELLER_IDS` map group → credential; legacy single vars (`SP_API_REFRESH_TOKEN`, `SELLER_ID`) still work as the EU/default. Missing tokens simply cause that region's marketplaces to be skipped.

**Multi-marketplace registry:** the `MARKETPLACES` dict maps each country code → `{marketplace_id, native lang, list-locale, currency, region, token-group}`. `ENGLISH_ONLY = True` sends English text, tagged with the local locale where a store rejects English item names (BE/NL/SE/PL/MX/JP). One `ListingsItems` client per **region endpoint** (`REGION_ENDPOINT`: EU→UK, NA→US, FE→JP) serves all that region's marketplaces; the actual target is passed per call via `marketplaceIds`. `set_active_marketplace(code)` repoints the module globals `MARKETPLACE_ID`, `LANGUAGE_TAG`, `CURRENCY`, `SELLER_ID` per row — safe because processing is single-threaded.

**Data flow (per poll, `process_sheet`):**
1. Each data row → `parse_row()` → typed product dict (its `countries` cell lists target marketplaces).
2. **Action column (O)** checked first — values in `DELETE_TRIGGERS` (`DELETE`/`TRUE`/`X`/`YES`/`1`/`REMOVE`) trigger deletion across that row's markets.
3. **Change detection, both directions:**
   - `sheet_content_hash()` vs the hash stored in column N → did the sheet change?
   - `amazon_detect_hash()` / `read_amazon_state()` vs column P + Amazon's `lastUpdatedDate` (column Q) → did Amazon change?
   - Per-market state (asin + content hash + last local price/currency) is packed as compact JSON in **column W**, so markets already live at the current content **and current local price** are skipped with no API call.
4. **Conflict rule — the SHEET is primary.** Only-sheet-changed → push to Amazon. Only-Amazon-changed → `pull_updates()` writes Amazon's values back into the sheet. **Both changed → sheet wins**, and Amazon's value + timestamp are recorded in the CONFLICT column (T) for manual review.
5. **Parent before child** — one non-purchasable variation parent (`build_parent_payload`, `requirements=LISTING_PRODUCT_ONLY`) per unique `parent_sku`, then each purchasable child (`build_child_payload`, `requirements=LISTING`). `put_listings_item` is an upsert (creates and updates).
6. **Verify-after-submit** (`VERIFY_AFTER_SUBMIT`): after a push, poll the listing a few times (`VERIFY_MAX_ATTEMPTS` × `VERIFY_INTERVAL_SECONDS`) to surface the real accepted/errored outcome instead of "queued".
7. Outcome written back to the sheet; column L shows a per-market summary (e.g. `OK:DE,UK | RETRY:JP`).

**Pricing — three steps, one entry point.** Column G is the **wholesale cost in USD** (`BASE_CURRENCY`), written by `fetch_esim_prices.py` from the supplier's live catalogue — not the shelf price. `market_price(cost, code)` is the only function anything calls, and it does:

0. **Cost → `PRICING_CURRENCY` (GBP)** via `convert_amount()`, because the policy below is written in pounds while the supplier quotes USD.
1. **`retail_price_raw()` — competitive tier markup**, in `PRICING_CURRENCY`:
   `<£5 → ×3` · `£5–£15 → ×2` · `£15–£40 → ×1.6` · `£40+ → ×1.3` (boundaries fall in the *cheaper* tier: exactly £15 is ×1.6). Then a **minimum-profit safeguard** — Amazon takes ~`AMAZON_FEE_PCT` (18%), so the price is lifted to `(cost + MIN_PROFIT) / (1 − fee)` whenever the multiple alone would clear less than `MIN_PROFIT` (£4). The higher of the two wins; the safeguard only ever raises. The result is **left unrounded** — see step 2. Cost $11.25 → £8.55 → ×2 = 17.10, floor 15.30 → **17.10 raw**.
   - The fee rate is a blended approximation of the **referral fee only** — no FBA fulfilment/storage fees or VAT — so `MIN_PROFIT` floors gross contribution, not true net. Raise it to cover fulfilment.
   - Tier boundaries are **not monotonic by design**: base £14.99 → £29.99 but base £15.00 → £24.99 (same at £40). That's the specified policy, not a bug.
   - `retail_price()` (no `_raw`) is that same number rounded up to the next `.99`: the **UK's own shelf price**, and the headline figure in previews and logs. Nothing else prices off it.
2. **`market_price()` — live FX, then the `.99`**, converting the *raw* GBP selling price into the marketplace's currency and only then rounding **up to the next .99 in that currency**. `convert_amount()` crosses any two currencies via the base, so USD→GBP→PLN needs no intermediate rounding.
   - **The order matters and was a real bug** (fixed 2026-08-13): rounding to `.99` in GBP *and* again after conversion stacked two round-ups. Cost $1.50 → raw £6.23 → old: £6.99 → €8.15 → **€8.99**; new: €7.27 → **€7.99**. Same for PL (35.99 zł → 31.99 zł), SE, US, CA, MX, JP, AU. The UK is unaffected (GBP is the pricing currency, so it rounds once either way).
   - Every store therefore shows a genuine `x.99` in **its own** currency, not the FX image of a British price point. JPY is the one exception — whole yen, so charm rounding uses the `…,980` / 100-yen grid.

**`fetch_esim_prices.py`** fills column G. It POSTs `api.esimaccess.com/api/v1/open/package/list` (header `RT-AccessCode`, creds in `.env`), caches the ~2,900-package catalogue to `esim_packages.json` for 6 h, and matches each sheet row **by name, not SKU**: both sides reduce to `(place, GB, days)` — `"Bulgaria eSIM 3GB 15Days | Pay As You Go"` and `"Bulgaria 3GB 15Days"` both → `("bulgaria", 3.0, 15)`. Accents are folded (Curaçao/Åland), marketing words and `(30+ areas)` stripped, `/Day` plans kept distinct from fixed bundles. **API prices are in 1/10000 USD** (112500 → $11.25) and `retailPrice` is exactly 2× `price` — column G takes `price`, the wholesale cost. Where several packages share a key the **cheapest** wins and the alternatives are logged. Unmatched rows are reported, never guessed: put a `packageCode` in **column X** to pin one (a valid pin is always honoured; stray non-code text in X is ignored). Previews by default; `--apply` writes.

`TIER_PRICING = False` reverts to "base price = selling price". Tunable from `.env`: `AMAZON_FEE_PCT`, `MIN_PROFIT`, `BASE_CURRENCY`. **Changing any of them reprices the catalogue on the next pass** — check `python upload_listings.py --rates <real base price>` first; it prints the whole tier + FX breakdown and writes nothing.

**Exchange rates.** `fx_rates()` fetches from three free key-less feeds in order (open.er-api.com → frankfurter.app → currency-api) and caches in memory **and** on disk (`fx_rates.json`, gitignored) for `FX_TTL_HOURS` (12), so a 5-minute poll refetches at most twice a day and a restart reuses the cache. Failure ladder: fresh memory → fresh disk → live fetch → **stale disk (loudly logged)** → market reported `NO_FX` and **skipped**. Never send an unconverted number under a foreign currency code — that is the one outcome the design refuses.
- **Rounding** (`FX_ROUNDING`, default `charm`) rounds **up** to the next `.99` **after** conversion — the only place the `.99` is applied for a non-GBP market — so a rate move can never undercut margin; zero-decimal currencies (JPY) use a `…,980` ending on a 100-yen grid. `FX_MARKUP_PCT` adds an uplift over the pure rate; `FX_RATE_<CUR>` in `.env` pins a currency manually.
- **The final local price is what column W's per-market hash is built from**, so a market re-pushes only when *its own rounded* price changes — with charm rounding a rate must drift ~1–3% before anything moves. Enabling tier pricing changes every market including the UK, so the first pass after it re-pushes the whole catalogue once.
- **Price is never pulled back from Amazon while `TIER_PRICING` is on** (`pull_updates`): column G is the cost input the formula reads, so mirroring a marked-up, FX-converted shelf price into it would re-mark-up on the next pass. All other reverse-sync fields still pull. With tier pricing off, price pulls again via `base_price_from()` (FX inverse only).

**Poll cadence:** activity-based backoff — `POLL_MIN_SECONDS` (300) after an active pass, growing by `POLL_BACKOFF_SECONDS` (150) each idle pass up to `POLL_MAX_SECONDS` (600).

**Column contract:** `COL_*` constants (0-based) map sheet columns to fields and **must match the sheet's header order**. Sheet rows are 1-based with row 1 = header, so `sheet_row = data_row_index + 2`.
- **A–J** user input: sku, parent_sku, parent_title, title, variation, brand, price (**wholesale cost in USD, not the shelf price** — written by `fetch_esim_prices.py`; tier markup + FX are applied on push), quantity, description, image1.
- **K–N** written by script: uploaded (K), status (L), asin (M), sheet-content hash (N).
- **O** user-controlled action.
- **P–T** reverse-sync bookkeeping written by script: amazon_hash (P), amazon_updated (Q), sheet_updated (R), last_synced (S), conflict (T).
- **U** user product_id (EAN/GTIN barcode; falls back to GTIN exemption if blank).
- **V** user **countries** list (single code / comma list / region / `ALL`).
- **W** written by script: per-market state JSON.
- **X** package_code — the eSIM Access `packageCode` a row is priced from. Written by `fetch_esim_prices.py`; fill it in by hand to pin a row.

**Status sentinels** in column L: `UPLOADED`/`UPDATED`/`PULLED` (live — see `LIVE_STATUSES`), `DELETED` (stays deleted until action cleared), `HOLD` (manually freeze a row — skipped). Re-running is idempotent thanks to the dual hash check.

**Attribute formatters** (`text_attr`, `enum_attr`, `text_attr_multi`) wrap values in the localized `{value, language_tag, marketplace_id}` structures SP-API requires. `compliance_attrs()` holds the SIM_CARD product-level attributes Amazon mandates on both parent and child.

**`fetch_listings.py`** deliberately uses the **Reports API** (`GET_MERCHANT_LISTINGS_ALL_DATA`), not `searchListingsItems`, because for this account searchListingsItems returns only one page with no pagination token. It writes to a separate tab (`amazon_current`) and never touches the uploader's tab.

**`fetch_orders.py`** exports orders (not listings) via the **Orders API** — one Excel workbook per marketplace into `Sheets/`, default SE + PL. Tab `orders` is one row per order *line* (order number, purchase date, name, SKU, ASIN, qty, price, status, fulfilment); tab `products` collapses that to unique name/SKU/ASIN with order + unit counts. Read-only.
- **Rate limits are the runtime**: getOrders is 0.0167 rps (burst 20), getOrderItems 0.5 rps (burst 30), and ASIN/SKU/name only exist on the *items* call — so ~300 orders is ~300 calls at 2 s, ~10 min per marketplace.
- **So every response is cached** to `reports_cache/orders_<CODE>.json` and flushed every 25 orders: a re-run is seconds, and an interrupted run resumes. The cache records the window it was fetched with, so asking for *more* history re-fetches the order list (item detail, keyed by order id, is kept). `--refresh` discards it.
- `--since` defaults to 2015-01-01 — i.e. "everything Amazon still has", since the Orders API only retains a couple of years; if a window is rejected as too old, `DATE_FALLBACKS` walks forward.
- **Pending orders carry no pricing** (Amazon withholds it until payment authorises) and a purged order returns no items at all. Both keep their row with the missing fields blank — the order number was the point — and item-less rows are greyed in the sheet. Cancelled orders are included unless `--exclude-cancelled`.

**`inspect_listing.py`** is a read-only debugging aid for the reverse sync: it prints one SKU's raw Amazon payload and exactly what `read_amazon_state()` extracts, so you can confirm reverse-sync field mapping before going live.

**`fbm_to_fba.py`** is a one-off migration tool: it reads the SKU list from `FBM to FBA/*.xlsx` and PATCHes each listing's fulfilment channel. It **previews by default** (`mode=VALIDATION_PREVIEW`) and only writes with `--live`; `--verify-only` audits current channels. It imports `fba_channel_for`/`FBM_CHANNEL` from `upload_listings` so the two can't disagree. Four live-verified SP-API behaviours are encoded in it — the module docstring has the detail, but in short:
- `fulfillment_availability` is **selector-keyed** on `fulfillment_channel_code`, so a `replace` op *merges* instead of replacing and leaves the SKU on both channels, still fulfilled by `DEFAULT`. The merchant entry must be deleted explicitly **by value** (a valueless `delete` is rejected).
- A converted SKU keeps a **residual `DEFAULT` entry in `attributes` forever**. The authoritative live state is the top-level `fulfillmentAvailability` (offer) view — verify against that, never the attributes list.
- The offer view **lags ~30–60s** behind a patch, so the script patches everything, waits once (`SETTLE_SECONDS`), then verifies. Per-SKU waits would turn ~30 min into ~7 h. `ACCEPTED` from a patch means *queued*, not applied — always verify.
- `VALIDATION_PREVIEW` validates a patch **in isolation** (no merge with the live listing), so it falsely reports required attributes missing (`90220 batteries_required`). `ECHO_ATTRS` echoes them back unchanged to satisfy it.

Converting the channel only moves the **offer**. It sends no stock: listings stay out-of-stock until an FBA inbound shipment is received, and the conversion cancels the merchant-fulfilled offer.

**`market_activation.py`** was written to answer "sell in the UK only" — **it cannot do that, and a live run took amazon.co.uk down on 2026-08-18. Read the `fbm0` bullet below before running it at all.** `REGION_COUNTRIES` lists every marketplace per region and `ACTIVE_MARKETS` (currently `["UK"]`) names the ones allowed to sell; everything else is **deactivated, never deleted** — the SKU, ASIN, content and history survive and can be switched back on. **The Google Sheet is the work list** (read only — no cell is written; `--from-xlsx` reads a downloaded copy offline): each row is expanded by its countries column, so one sheet row is a *separate product in every country it names* — 457 SKUs × their countries, of which 4,949 are switched off one by one. The cached merchant-listings report `purge_listings.py` uses supplies Amazon's side (live price, status, ASIN) and adds anything live that the sheet does not know about, so rogue old-naming listings are covered too. Output is one local `market_activation_<stamp>.xlsx`: ~7,670 rows of country, sku, name, ASIN, then **the pricing chain one column per step** (`price_chain()`: eSIM Access price USD → formula price GBP after tier+min-profit → converted at today's rate → final `.99` shelf price → currency), Amazon's live price, a GREEN/RED verdict that names the gap in the shopper's own currency (`TOO LOW by 36.00 PLN`), a GREEN/RED status verdict, and a `source` column (SHEET+AMAZON / SHEET only / AMAZON only), plus a per-marketplace summary tab.
- **Column W is read as a second source.** `parse_market_state()` unpacks the sync loop's per-market JSON (`{"BE":{"asin":…,"price":31.99},"DE":{"err":"…untergeordneten ASIN…"}}`) to fill in the ASIN and price where the listings report has no row — and it is the only place the per-market **failure reason** lives, so a market with no listing says *why* in the Notes column instead of just reading blank. Unparseable cells are ignored (older sheets held `generated_from` in W). Previews by default; `--live` applies; `--limit N` is a canary.
- **`--method fbm0` IS NOT PER-MARKETPLACE — it switches the SKU off across the whole region, UK included.** `fulfillment_availability` entries are `{fulfillment_channel_code, quantity}` with **no `marketplace_id` selector**, so the attribute is SKU-level for the entire SP-API region. Patching it with `marketplaceIds=[<IE>]` changes the offer on **every EU marketplace, including amazon.co.uk**. Proven live 2026-08-18 by reading one SKU (`AALAND-ISLANDS-3GB-15D`) in three marketplaces — all three returned the identical offer `[('DEFAULT', 0)]`, status `['DISCOVERABLE']`.
  - **What it cost.** The 2026-08-18 run deactivated 9 EU marketplaces and collaterally switched off the UK: active UK listings fell **119 → 23**, and **657** UK non-parent listings landed on `DEFAULT` at quantity 0 — 657 of 657 being SKUs the script had patched in some *other* marketplace. The ledger correctly holds zero UK rows; the damage was entirely collateral. Restored with `fbm_to_fba.py --markets UK --excel <list> --live`, which is region-wide in the same way and therefore turned the EU offers back on too.
  - **So "UK selling, EU off" is unreachable through `fbm0`** — both markets share the single attribute being changed. A genuine per-marketplace method still has to be found; until then this script's core purpose is unmet.
  - The earlier claim "verified live on amazon.ie" was **true but insufficient**: it confirmed the IE offer went `DISCOVERABLE,BUYABLE` → `DISCOVERABLE` and never checked the protected marketplace. **Verify any future method by reading the same SKU back in the marketplace that must KEEP selling, not only in the one being switched off.**
  - Deactivation is still a PATCH, never `deleteListingsItem` — no SKU, ASIN, content or history is ever destroyed, and the channel move is reversible.
  - **Allow ~5 minutes before judging a patch.** The offer view lags far longer than the 30–60 s `fbm_to_fba.py` documents: at +4 min the attributes still listed `AMAZON_EU` and the SKU still read BUYABLE; at +6 min it had flipped. `lastUpdatedDate` is **not** a usable signal — it stayed at 13 Aug on SKUs whose attributes had demonstrably changed.
  - **`--method offer` does not work — verified live, don't reach for it.** Deleting `purchasable_offer` was the original default. Tested on amazon.de 2026-08-18 against three AALAND-ISLANDS SKUs: Amazon returned ACCEPTED, the attribute really was gone (`attributes.purchasable_offer` = null), *and the listing stayed `BUYABLE, DISCOVERABLE` selling at €8.99 eight minutes later*. The offer keeps its price independently of the attribute. The flag is kept only so the finding isn't rediscovered the hard way.
- **The UK guard works, but guards the wrong thing.** `NEVER_DEACTIVATE = {"UK"}` and `is_protected()` gate the marketplace loop *and* `deactivate_one()`, so no call is ever *aimed at* amazon.co.uk — and indeed the ledger has never held a UK row. **That is not the same as the UK being safe**, because the attribute being patched is region-wide (above): on 2026-08-18 the UK went down without a single UK call. A guard on the call's target marketplace cannot protect against an attribute whose scope is the whole region — keeping the UK selling is the client's one hard requirement, so **prove protection by reading the UK back after a canary, never by inspecting the guard**.
- **Every product is attempted regardless of its reported status** — a report is up to a day old and "Inactive" usually means out-of-stock, not withdrawn, so it would return the moment stock landed. `--skip-inactive` restores the cheaper behaviour for re-runs. Sheet products the report shows as never listed in a marketplace are reported, not called; `--verify-unlisted` asks Amazon anyway.
- Variation **parents are never touched** (no offer to switch off, and it would dissolve the family). Live changes are journalled to `logs/deactivate_ledger.csv` and skipped on re-run, so an interrupted run resumes.
- The Excel is rewritten **after every marketplace** and the ledger flushed every `LEDGER_FLUSH` rows, so a two-hour run has a usable file on disk throughout instead of nothing until it ends.
- **It does not stop the sync loop re-listing.** `upload_listings.py` pushes every marketplace named in column V, so deactivated markets come back on the next pass unless column V is narrowed to `UK` (or the loop is stopped).
- Only the **EU** authorization group has a refresh token in `.env`, so US/CA/MX/JP/AU cannot be read or changed — they are reported as skipped, not as "clean".

## Domain constraints (already encoded — don't "fix" without reason)

- **Product type is `SIM_CARD`**, not WIRELESS_ACCESSORY — confirmed against the live Category Listings Report. The `sim_card_schema_*.json` files are the authoritative attribute reference; regenerate with `fetch_schema.py`.
- **Fulfilment is FBA** (`USE_FBA = True`, top of `upload_listings.py`): YouGuide prints physical cards and stocks them at Amazon. `fba_channel_for(code)` resolves the network per marketplace (`AMAZON_EU` on UK/EU, `AMAZON_NA`, `AMAZON_JP`, `AMAZON_AU`) — key it by **marketplace, not region**: JP and AU share the FE endpoint but are different networks. Under FBA the sheet's quantity (col H) is **not** sent — Amazon owns the quantity, derived from units received at a fulfilment centre. Set `USE_FBA = False` to go back to `DEFAULT` (Merchant Fulfilled).
  - This reverses an earlier constraint that eSIMs "cannot be FBA (warning 12998)". Re-verified against the live API on 2026-07-17: the UK SIM_CARD schema enumerates `AMAZON_EU` as valid and Amazon accepted it with no 12998 warning.
- GTIN exemption is claimed when no barcode (col U) is supplied. Variation theme is `SIZE` (SIM_CARD deprecated STYLE_NAME); the sheet's variation label is written to the `size` attribute.
- Listing language/currency **follow the marketplace** (per the `MARKETPLACES` registry); text is sent as-is (no translation), but the **price is marked up by tier and converted at the live rate** — see "Pricing — two steps, one entry point".
- Throttle sleeps (`SLEEP_AFTER_*`) respect SP-API and Google Sheets quotas — don't remove them.

## Deployment

`deploy/DEPLOY.md` covers running the loop on a DigitalOcean droplet via systemd (`deploy/youguide-sync.service`). Key gotcha: **do the Google OAuth once on a machine with a browser and copy `authorized_user.json` to the server** — the droplet has no browser for first-run auth. Files that must travel to the server: `upload_listings.py`, `.env`, `oauth_client.json`, `authorized_user.json`.

## Other assets

`YouGuide_GPSR_Amazon_Pack/` holds EU GPSR safety-compliance documents (multilingual PDFs + images) for manual upload to Amazon — reference material, **not** consumed by any script. `Amazon Listings.xlsx` and the `.backup` are working copies of the sheet.
