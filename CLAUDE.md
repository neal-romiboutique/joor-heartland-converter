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

- **Size completion** (`detect_convention` / `complete_sizes`) is the least obvious part. A grid does not carry only the sizes ordered — it spans the union of the ordered range and a store-standard base range, per convention: letters `XS–L`, numeric-even `0–10` (with `"00"` a distinct rung *below* `0`, kept as a string throughout), denim `24–29`, euro `70–90`. Convention `"unknown"` sets `Generate Items = false` so an empty grid is created rather than wrong items.
- **Rings** are a separate ladder (`5–9`, base `6–8`). Ring and apparel size tokens are identical, so `detect_convention` can't tell `6,8` apart on its own — it takes an `is_ring` flag that `build_grids` derives from the style name via `looks_like_ring` (`\brings?\b`, which deliberately does not match `Earrings`/`Herringbone`). The ring branch **must stay ahead of the numeric branches** or `numeric-even` claims all-even ring orders first. When a vendor encodes the size in the style number (Jenny Bird `JB4014-6-HPG-W`), each size is its own grid, so `_size_encoded_in_style_number` suppresses ladder completion — otherwise every sibling grid would repeat the whole base run.
- `expand_size_ranges` deliberately does **not** expand: a token like `00-14` (Veronica Beard) is one literal size, not a run.
- **Pricing** — JOOR's Sugg. Retail when present, otherwise `MARGIN` (2.4) × cost, always ceiled to whole dollars. Note that JOOR's Order-to-Size export writes `WholeSale` pre-rounded to whole dollars: a vendor pricing at retail ÷ 2.2 (Jenny Bird) exports 108 → 50 while their UI displays 49.00. The cost is lossy on arrival, and `ceil_dollars` is not the cause — don't "fix" it there.
- **Colors** (`normalize_color`) title-case and join multi-color names with ampersands (`BLUE/BLACK` → `Blue & Black`), protecting `W/` as shorthand for "with". Grids created before this change are stored with slashes, so `_color_variants` retries the legacy slash form on dedupe and item lookup.
- **Vendor resolution** (`normalize_vendor`) matches against the hardcoded `KNOWN_VENDOR_NAMES` list via an accent/case/punctuation-stripping key. The point is to hit an *exact existing* Heartland vendor name — a near miss creates a duplicate vendor record on import. Unmatched vendors warn and fall through to a guess. Regenerate the list when vendors are added in Heartland; use `VENDOR_MAP` for one-off overrides.
- `TEMPLATE_HEADERS` (84 columns) is Heartland's official template verbatim — do not reorder, rename, or trim it. `DROPDOWN_LISTS` columns render as in-cell dropdowns via a hidden `Lists` sheet so the human reviewer picks controlled values instead of typing.
- `PO #` is written as the literal placeholder `CHECK HEARTLAND HIGHEST PO#`: Heartland requires the field on import but won't auto-number, so the user fixes it manually. The real order number goes in `PO Description`.

### Heartland API

Credentials come from `.env` next to the script (`HEARTLAND_SUBDOMAIN`, `HEARTLAND_TOKEN`), loaded by a hand-rolled parser — no python-dotenv. Base URL is `https://{subdomain}.retail.heartland.us/api`. All calls are read-only GETs. Reference docs: https://dev.retail.heartland.us/

Two endpoint quirks are easy to get wrong:

- `/api/items` filters custom fields with the plain `~[custom@<field>]` prefix, and the vendor with `~[primary_vendor_id]`. *(Confirmed against Heartland's published docs.)*
- `/api/item_grids` filters with `~[item_custom@<field>]` and `~[item_primary_vendor_id]` — note the extra `item_` prefix on both. *(Established empirically by `api_testing/api_test_5.py` after `custom@` returned HTTP 400; not stated in the public docs.)*
- `ITEM_NUMBER_FIELD = "id"`, not `public_id`. Vendors proved these differ; the constant exists as a single switch if `api_test_8.py` ever shows otherwise.

The `api_testing/api_test_*.py` scripts are the record of how these were determined — each one's docstring states the question it answered. They are read-only probes, safe to re-run.

`build_po` fetches one paginated index of all of a vendor's items rather than one request per order line, and clears `_ITEM_INDEX_CACHE` at the start of each run because the grids import likely just created new items. API failures return `None` and are treated as "include the grid / stop the PO" — never as "silently drop the line".

### vendor_id_map.py

`order_to_heartland.py` imports `VENDOR_ID_MAP` from a sibling `vendor_id_map.py`, which is **gitignored** and currently only exists in `api_testing/`. Without it at the repo root, every live dedupe and item lookup is skipped: grids are always generated (never checked against existing ones) and `build_po` exits with an error. `api_test_6.py` writes it into `api_testing/`, so it must be copied or symlinked up to the root.

## Conventions to preserve

- User-facing failures use `sys.exit("ERROR: ...")` with a sentence explaining what to do next. `converter_app.run_conversion` catches `SystemExit` and shows the message in the browser log, so these must stay human-readable — the shop owner reads them, not a developer.
- Progress goes through `print()`, and the frontend regex-styles lines starting with `WARNING`, `NOTE`, and `wrote ` (amber, amber, green). Keep those prefixes.
- The comments explaining *why* a rule exists (which vendor's export forced it, which API probe confirmed a field name) are the main documentation of the Heartland integration. Extend them rather than trimming them.

## Local-only files

`.env`, `*.xlsx`, and `vendor_id_map.py` are gitignored — order exports and generated import files contain live vendor pricing. `Pull_Updates.command` exists so the shop owner can `git pull` by double-clicking.
