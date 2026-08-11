#!/usr/bin/env python3
"""Test 2: resolve a vendor by name, then fetch ONE item grid for that vendor
to inspect its raw JSON (field names + custom field keys). Read-only."""
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

VENDOR_NAME = "Veronica Beard"   # change this if you want to test a different vendor

# --- Step 1: resolve vendor name -> id ---
resp = requests.get(
    f"{BASE_URL}/purchasing/vendors",
    headers=HEADERS,
    params={"query": VENDOR_NAME, "per_page": 5},
)
print("Vendor lookup status:", resp.status_code)
data = resp.json()
results = data.get("results", [])
if not results:
    print(f"No vendor found matching '{VENDOR_NAME}'. Raw response:")
    print(resp.text[:500])
    raise SystemExit(1)

vendor = results[0]
vendor_id = vendor["id"]
print(f"Matched vendor: {vendor['name']} (id={vendor_id})")

# --- Step 2: fetch one item grid for that vendor ---
resp2 = requests.get(
    f"{BASE_URL}/item_grids",
    headers=HEADERS,
    params={
        "~[item_primary_vendor_id]": vendor_id,
        "per_page": 1,
    },
)
print("\nItem grid lookup status:", resp2.status_code)
print(resp2.text[:3000])
