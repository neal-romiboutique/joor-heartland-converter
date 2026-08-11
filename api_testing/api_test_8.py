#!/usr/bin/env python3
"""Test 8: determine which field on /api/items corresponds to the "Item #"
that Heartland's UI displays and its PO import expects -- `id` (internal)
or `public_id` (derived/displayed), mirroring the vendor id/public_id trap.
build_po currently uses `id`; if that's wrong, every PO import file carries
plausible-looking but incorrect Item #s. Read-only."""
import os
import requests
from pathlib import Path

env_path = Path(__file__).resolve().parent / ".env"
for line in env_path.read_text().splitlines():
    if line.strip() and not line.startswith("#") and "=" in line:
        k, v = line.split("=", 1)
        os.environ.setdefault(k.strip(), v.strip())

TOKEN = os.environ["HEARTLAND_TOKEN"]
SUBDOMAIN = os.environ["HEARTLAND_SUBDOMAIN"]
BASE_URL = f"https://{SUBDOMAIN}.retail.heartland.us/api"
HEADERS = {"Authorization": f"Bearer {TOKEN}"}

# ---- FILL THESE IN from an item you can see in the Heartland UI ----
VENDOR_ID    = 100387              # e.g. 100387 = Veronica Beard
STYLE_NUMBER = "27RE0168"
COLOR        = "Navy Multi"
SIZE         = "2"                  # optional; leave "" to match any size
UI_ITEM_NUM  = "203326"     # the Item # shown in Heartland's UI
# --------------------------------------------------------------------

params = {
    "~[primary_vendor_id]": VENDOR_ID,
    "~[custom@style_number]": STYLE_NUMBER,
    "~[custom@color]": COLOR,
    "per_page": 3,
}
if SIZE:
    params["~[custom@size]"] = SIZE

resp = requests.get(f"{BASE_URL}/items", headers=HEADERS, params=params)
print("status:", resp.status_code)
if resp.status_code != 200:
    print(resp.text[:800])
    raise SystemExit(1)

results = resp.json().get("results", [])
print(f"total matched: {resp.json().get('total')}  returned: {len(results)}\n")
if not results:
    raise SystemExit("No items matched -- check vendor id / style / color.")

# --- full field dump of the first match (find any other #-looking fields) ---
first = results[0]
print("--- ALL top-level fields on first match ---")
for k, v in sorted(first.items()):
    if not isinstance(v, (dict, list)):
        print(f"  {k}: {v!r}")
print("\n--- nested/custom blocks ---")
for k, v in first.items():
    if isinstance(v, (dict, list)):
        print(f"  {k}: {str(v)[:300]}")

# --- the actual comparison ---
print("\n--- COMPARISON vs Heartland UI ---")
print(f"  UI shows Item #:      {UI_ITEM_NUM!r}")
print(f"  API 'id':             {first.get('id')!r}")
print(f"  API 'public_id':      {first.get('public_id')!r}")
for k in ("number", "item_number", "sku", "code"):
    if k in first:
        print(f"  API {k!r}: {first[k]!r}")

print("\n  -> whichever field EQUALS the UI value is what build_po must use.")

# --- sanity: show all matches so size-level distinctness is visible ---
print("\n--- all returned matches (id / public_id / size / description) ---")
for it in results:
    cust = it.get("custom") or {}
    print(f"  id={it.get('id')!r}  public_id={it.get('public_id')!r}  "
          f"size={cust.get('size')!r}  desc={str(it.get('description'))[:40]!r}")
