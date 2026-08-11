#!/usr/bin/env python3
"""Test 6: page through ALL vendors and print id / public_id / name for each,
so we can build an accurate name->id lookup dict (and see how public_id
behaves for vendors with special characters like '&'). Read-only."""
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

all_vendors = []
page = 1
while True:
    resp = requests.get(
        f"{BASE_URL}/purchasing/vendors",
        headers=HEADERS,
        params={"per_page": 100, "page": page},
    )
    if resp.status_code != 200:
        print(f"ERROR on page {page}: {resp.status_code} {resp.text[:300]}")
        break
    data = resp.json()
    results = data.get("results", [])
    all_vendors.extend(results)
    print(f"page {page}/{data.get('pages')}: got {len(results)} (running total {len(all_vendors)})")
    if page >= data.get("pages", 1):
        break
    page += 1

print(f"\nTotal vendors fetched: {len(all_vendors)}\n")

# Show a few examples where public_id looks numeric (like "Lola & Sophie")
print("--- Vendors where public_id is purely numeric (edge cases) ---")
for v in all_vendors:
    pid = str(v.get("public_id") or "")
    if pid.isdigit():
        print(f"  name={v['name']!r}  id={v['id']}  public_id={v['public_id']!r}")

# Write the full name -> id dict to a local file (not printed in full to keep
# terminal output manageable)
out_path = Path(__file__).resolve().parent / "vendor_id_map.py"
with open(out_path, "w") as f:
    f.write("# Auto-generated from Heartland /api/purchasing/vendors\n")
    f.write("VENDOR_ID_MAP = {\n")
    for v in sorted(all_vendors, key=lambda x: x["name"].lower()):
        name = v["name"].replace("'", "\\'")
        f.write(f"    '{name}': {v['id']},\n")
    f.write("}\n")

print(f"\nWrote full name->id dict for {len(all_vendors)} vendors to: {out_path}")
