# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A single-store internal tool for Romi Boutique that converts wholesale order exports (JOOR and NuOrder) into the Excel import files Heartland Retail expects. The end user is a non-technical shop owner who double-clicks a `.command` file; the code has no build step, no test suite, and no dependency manifest.

## Commands

```bash
# Install deps (no requirements.txt exists)
python3 -m pip install openpyxl requests

# Run the drag-and-drop app (opens a browser tab on a random localhost port)
python3 converter_app.py           # or double-click "Order Converter.command"

# Run the converter directly
python3 order_to_heartland.py grids ORDER.xlsx -o grids_import.xlsx
python3 order_to_heartland.py po    ORDER.xlsx -o po_import.xlsx

# Regenerate vendor_id_map.py from live Heartland (see "vendor_id_map.py" below)
python3 api_testing/api_test_6.py
```

There is no lint/test/build. Verification is manual: run a real order export through and read the printed log.

## Architecture

Two files do everything:

- **`order_to_heartland.py`** — all parsing, business rules, Heartland API calls, and xlsx writing.
- **`converter_app.py`** — a stdlib `ThreadingHTTPServer` on `127.0.0.1` serving one hardcoded HTML page (`PAGE`) with three tabs. It base64s the uploaded file over `POST /convert`, calls into the converter with `stdout` redirected into a string, and returns `{log, files:[{name, b64}]}` which the page turns into browser downloads. It contains **zero** conversion logic — keep it that way.

Docstrings and the import-error message still refer to `joor_to_heartland.py` / `joor_converter_app.py`; the files were renamed in commit `64d407a` and the text is stale.

### The long-row dict is the contract

Every reader normalizes its input into a list of the same flat dicts, and every builder consumes only those:

```
seller_account, style_number, style_name, color, size, quantity,
item_price, sugg_retail, order_number, start_date, cancel_date, SEASON_FROM
```

Readers: `read_joor` (auto-detects JOOR Basic Export long CSV/TSV/XLSX vs. `parse_order_to_size` wide XLSX, which melts size columns into rows) and `read_nuorder` (NuOrder "Order Data" tab; long already, but the vendor is *never* in the file so the caller must supply it — that is why the NuOrder tab shows a vendor text field). Builders: `build_grids` and `build_po`. Add a new source by writing a reader that emits these dicts, not by touching the builders.

Both `read_joor` branches return a plain list. They used to differ — the long path returned `(rows, {})` while the wide path returned a bare list — so the CLI crashed on Order-to-Size exports and the app crashed on Basic Exports. Keep the return shapes identical.

### Two-phase Heartland workflow

The phases are separated because Heartland must run in between them:

1. **grids** — emit one row per (vendor, style number, color) in Heartland's grid-import template. The user imports it under Inventory → Item Grids → Import, and Heartland generates one item per size.
2. **po** — re-read *the same* order export, look up each generated item's `Item #` live via the API, and emit the PO import file. Running `po` before the grids import has completed produces "no matching Heartland item" warnings for every line.

NuOrder is grids-only by design; its POs are built by hand in Heartland.

### Business rules encoded in the converter

- **Size completion** (`detect_convention` / `complete_sizes`) is the least obvious part. A grid does not carry only the sizes ordered — it spans the union of the ordered range and a store-standard base range, per convention: letters `XS–L`, letters-petite `P–L`, combo-letters `P–M/L`, numeric-even `0–10` (with `"00"` a distinct rung *below* `0`, kept as a string throughout), denim `24–29`, euro `70–90`. Convention `"unknown"` sets `Generate Items = false` so an empty grid is created rather than wrong items — and because such a grid has no items, dedupe reads it as absent and re-emits it on every subsequent run until the sizes are added by hand.
- **Petite is its own ladder**, not a rung on `LETTER_LADDER`. 98 grids in Heartland carry `{P,S,M,L}` (85 of them Ulla Johnson) and none carry `XS`, so adding `P` to the main ladder would put a Petite size in every ordinary letter grid, since `LETTER_BASE` spans `XS–L`. `detect_convention`'s petite branch is gated on `P` actually being ordered for exactly that reason. `{P, XS}` together still falls to `unknown` (2 grids storewide) — the safe direction.
- **Combined-letter sizing** (`COMBO_LADDER`, TWP and A Shirt Thing) treats `XS/S` as one size token, not a slash-joined color. Only the rungs the store actually uses are on it (`P, XS/S, M/L, 1XL`); inventing a plausible `S/M` between them would change the shape of all 51 existing grids. Detection requires a slashed token so a plain `{P}` order stays on the petite ladder.
- **Small consecutive numeric ladders** (`numeric-ordered`, Lisa Yang `00,0,1,2,3`) get **no completion** — the ordered sizes are emitted as-is, in numeric order. There is no store-standard base run to complete to, and the branch this replaced guessed one: an order of `0,1,00` came back from Heartland as `0,1,2,3,4`, with three sizes never ordered and the `00` rung lost to `int("00") == 0`. `numeric-even` is checked first, so plain even runs still complete to `0–10`.
- **Size ORDER matters and is not sorted.** Every ladder branch slices the ladder, so order is correct by construction. The `unknown` fallback used to return `sorted(set(...))`, which is alphabetical — that is what put Ulla Johnson's `P,S,M` into Heartland as `M,P,S`. It now sorts through `_size_sort_key`, which knows the letter ladders, `P`, `00`, and the combo tokens.
- **Rings** are a separate ladder (`5–9`, base `6–8`). Ring and apparel size tokens are identical, so `detect_convention` can't tell `6,8` apart on its own — it takes an `is_ring` flag that `build_grids` derives from the style name via `looks_like_ring` (`\brings?\b`, which deliberately does not match `Earrings`/`Herringbone`). The ring branch **must stay ahead of the numeric branches** or `numeric-even` claims all-even ring orders first. When a vendor encodes the size in the style number (Jenny Bird `JB4014-6-HPG-W`), each size is its own grid, so `_size_encoded_in_style_number` suppresses ladder completion — otherwise every sibling grid would repeat the whole base run.
- `expand_size_ranges` deliberately does **not** expand: a token like `00-14` (Veronica Beard) is one literal size, not a run.
- **Pricing** — JOOR's Sugg. Retail when present, otherwise `MARGIN` (2.4) × cost, always ceiled to whole dollars. Note that JOOR's Order-to-Size export writes `WholeSale` pre-rounded to whole dollars: a vendor pricing at retail ÷ 2.2 (Jenny Bird) exports 108 → 50 while their UI displays 49.00. The cost is lossy on arrival, and `ceil_dollars` is not the cause — don't "fix" it there.
- **Colors** (`normalize_color`) title-case and join multi-color names with a bare **slash** (`BLUE/BLACK` → `Blue/Black`), protecting `W/` as shorthand for "with". The slash is the house standard: of 17,573 grids in Heartland, 1,358 carry a `/` in the color and only 40 carry an `&` — and most of those 40 are genuine vendor color *names* ("Sweet & Sassy"), not separators. Between 2026-08-21 and 2026-09-15 this function wrote ` & ` instead; those grids are matched, not rewritten.
- **Comparing two colors always goes through `_color_match_key`**, never `==`. It collapses `\s*[&/]\s*` to `/` and uppercases, so `Black/Cream`, `BLACK / CREAM` and `Black & Cream` are one color. Enumerating variants (the old `_color_variants`) missed the *spaced* slash form, which is a real stored form — Generation Love `H22369` has two grids differing only by the spaces. Measured: on one Veronica Beard order this took the PO's unmatched lines from 9 to 2 (the 7 were all `Black & Limestone` failing to match a stored `BLACK / LIMESTONE`). A genuine ampersand in a color name canonicalizes identically on both sides, so it still only matches itself.
- **Vendor resolution** (`normalize_vendor`) matches against `_VENDOR_LOOKUP` via an accent/case/punctuation-stripping key (`_vendor_key`). The point is to hit an *exact existing* Heartland vendor name — a near miss creates a duplicate vendor record on import **and silently disables dedupe**, because `resolve_vendor_id` then returns `None`. Order of resolution: `VENDOR_MAP` (keyed through `_vendor_key`, so casing doesn't matter) → exact key → `_vendor_near_matches` → guess + warning.
- `_vendor_near_matches` accepts names differing only by **trailing whole words** (JOOR "Mother Denim" → Heartland "Mother"; "AMO Denim" → "Amo") and auto-applies only when there is **exactly one** candidate, printing a NOTE. It must never pick between several: the live list holds real, distinct vendors separated by one trailing word — `Cami`/`Cami NYC`, `Milly`/`Milly Cabana`, `Jakett`/`Jakett New York`, `Frank & Eileen`/`Frank & Eileen Tee Lab`. Ambiguity warns and lists the candidates instead. Whole tokens only, so `Cami` can never match `Camilla`.
- **`build_grids` takes an optional `vendor_override`**, surfaced as the Vendor override box on the app's grids tab. It forces every grid onto one exact Heartland name and is the escape hatch when a JOOR brand name can't be resolved. NuOrder passes the same field (there it is required, since the export carries no vendor at all).
- `TEMPLATE_HEADERS` (84 columns) is Heartland's official template verbatim — do not reorder, rename, or trim it. `DROPDOWN_LISTS` columns render as in-cell dropdowns via a hidden `Lists` sheet so the human reviewer picks controlled values instead of typing.
- `PO #` is written as the literal placeholder `CHECK HEARTLAND HIGHEST PO#`: Heartland requires the field on import but won't auto-number, so the user fixes it manually. The real order number goes in `PO Description`.

### Heartland API

Credentials come from `.env` next to the script (`HEARTLAND_SUBDOMAIN`, `HEARTLAND_TOKEN`), loaded by a hand-rolled parser — no python-dotenv. Base URL is `https://{subdomain}.retail.heartland.us/api`. All calls are read-only GETs. Reference docs: https://dev.retail.heartland.us/

Two endpoint quirks are easy to get wrong:

- `/api/items` filters custom fields with the plain `~[custom@<field>]` prefix, and the vendor with `~[primary_vendor_id]`. *(Confirmed against Heartland's published docs.)*
- `/api/item_grids` filters with `~[item_custom@<field>]` and `~[item_primary_vendor_id]` — note the extra `item_` prefix on both. *(Established empirically by `api_testing/api_test_5.py` after `custom@` returned HTTP 400; not stated in the public docs.)* **Nothing in the converter filters grids by color any more — see "Dedupe asks about items, not grids" below.**
- Custom-field filters are **case-sensitive exact matches**. `~[custom@color]=BLACK` returns 4 items where `Black` returns 0, and stored casing is inconsistent within a single style (`CJ3548` carries both `BLACK` and `Charcoal`). Never rely on a server-side color filter; match client-side on an uppercased key.
- `ITEM_NUMBER_FIELD = "id"`, not `public_id`. Vendors proved these differ; the constant exists as a single switch if `api_test_8.py` ever shows otherwise.

The `api_testing/api_test_*.py` scripts are the record of how these were determined — each one's docstring states the question it answered. They are read-only probes, safe to re-run.

`build_po` fetches one paginated index of all of a vendor's items rather than one request per order line, and clears `_ITEM_INDEX_CACHE` at the start of each run because the grids import likely just created new items. API failures return `None` and are treated as "include the grid / stop the PO" — never as "silently drop the line".

### Dedupe asks about items, not grids

`_vendor_style_colors` reads `heartland_vendor_item_index` (the same per-vendor index `build_po` uses) and returns `{STYLE: {COLOR_KEY, ...}}` — "does any item exist for this (style, color)". It does **not** query `/api/item_grids`, because a grid's own `item_custom.color` is blank by design: `build_grids` writes the color to `Dimension 1 Values`, since a grid spans *all* colors of a style and has no single one. The original implementation filtered grids on that always-blank field, so it returned `False` for every grid this converter creates and dedupe never once fired. Measured before the fix: 0 of the 25 most recently imported grids had a grid-level color, and 6 of 411 sampled grids created since 2025.

Do not "simplify" this back to a style-only grid query. Four separate `RCK8232` grids were imported on 2026-08-25, one per color, all with blank grid-level color — matching on style alone would skip three real colors, and a missing item is far worse than a duplicate grid. Matching through the item index also makes the comparison case-insensitive for free (both sides are uppercased), which a server-side filter cannot be.

`build_grids` calls `clear_item_index_cache()` at the start of each run for the same reason `build_po` does: `converter_app.py` is a long-running server, so an index cached before the user's Heartland import would miss the items it just created.

**Whether the check ran is always stated in the log.** `build_grids` treats "couldn't check" as "emit the grid" — a visible duplicate the reviewer can delete beats a silently missing item — but two of the three ways to get there used to print *nothing at all*: a vendor name Heartland doesn't recognise, and a missing `.env`. A run with the check entirely disabled was indistinguishable from a run where every grid was genuinely new, which is how duplicate grids reached Heartland unnoticed. `dedupe_index_for_vendor` now prints one line per vendor either way, and the run ends with a `duplicate check: N skipped, N checked and new, N emitted UNCHECKED` summary. Keep that property when touching this code.

The index is built **once per vendor before the grid loop**, not per grid: `_vendor_style_colors` does not cache a failed fetch, so a per-grid call re-ran a full paginated request for every grid whenever Heartland was unreachable.

When a style exists for the vendor but not in this color, `build_grids` prints a NOTE listing the colors it *does* have. Usually that's a genuine new colorway, but it is also exactly what a color spelled differently looks like — it fired on a real ASKK NY order (`Berkeley` vs the stored `BERKLEY`) and would have caught the Generation Love `BLACK/CREAM` vs `BLACK / CREAM` duplicate.

Measured on the live catalog (2026-09-15): 162 duplicate `(vendor, style, color)` grid groups out of 17,573 grids, spread across the whole ID range — including Chan Luu grids 22760–22767 re-emitted wholesale as 24306–24317. All of them predate the dedupe fix, which is what pre-fix behavior predicts, since dedupe returned `False` for everything and any re-run duplicated the entire order.

### The vendor list refreshes itself

`refresh_vendor_lists()` runs at the top of `build_grids` and `build_po` and pulls every vendor from `/api/purchasing/vendors` (~450 vendors, ~3s), rebuilding **both** `_VENDOR_LOOKUP` (exact name for the import file) and `_VENDOR_ID_LOOKUP` (internal id for API calls) from that one response via `_rebuild_vendor_lookups`. New vendors appear in Heartland weekly and a stale list is expensive in both directions: an unknown name gets a guessed spelling that creates a duplicate vendor record on import, and an unmapped id silently disables dedupe.

Three properties worth preserving:

- **In memory only.** `vendor_id_map.py` is never rewritten, so every clone reads the live list straight from the API and no machine needs write access to the repo — which also keeps the shop machine git-clean so `Pull_Updates.command` can't hit a merge conflict.
- **Throttled by wall clock**, not a once-per-process flag (`VENDOR_REFRESH_SECONDS`, 24h). `converter_app.py` is a long-running server; a per-process flag would refresh at first use and then go stale for as long as the window stays open.
- **The clock only advances on success**, so a dropped connection retries on the next run instead of pinning the stale list for a day.

Any failure (offline, bad token, empty response) falls back to the committed snapshot and prints a NOTE — an empty API response must never wipe the list. `KNOWN_VENDOR_NAMES` and `vendor_id_map.py` are now only that offline baseline; they are allowed to age, and `api_testing/api_test_6.py` regenerates the latter at the repo root when you want it fresher. Live refresh requires `.env` on that machine — a clone without one silently runs on the snapshot.

## Conventions to preserve

- User-facing failures use `sys.exit("ERROR: ...")` with a sentence explaining what to do next. `converter_app.run_conversion` catches `SystemExit` and shows the message in the browser log, so these must stay human-readable — the shop owner reads them, not a developer.
- Progress goes through `print()`, and the frontend regex-styles lines starting with `WARNING`, `NOTE`, and `wrote ` (amber, amber, green). Keep those prefixes.
- The comments explaining *why* a rule exists (which vendor's export forced it, which API probe confirmed a field name) are the main documentation of the Heartland integration. Extend them rather than trimming them.

## Local-only files

`.env`, `*.xlsx`, and `vendor_id_map.py` are gitignored — order exports and generated import files contain live vendor pricing. `Pull_Updates.command` exists so the shop owner can `git pull` by double-clicking.
