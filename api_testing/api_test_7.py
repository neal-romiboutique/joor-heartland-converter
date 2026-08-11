#!/usr/bin/env python3
"""Test 7: confirm the custom-field prefix on /api/items (Step 2 needs this
to look up Item # by style_number/color/size, replacing the manual items
export). item_grids used "item_custom@"; items may use plain "custom@" per
the docs -- confirm rather than assume. Read-only."""
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

VENDOR_ID = 100387  # Veronica Beard
STYLE_NUMBER = "2109CD00210031"  # known grid from earlier tests
COLOR = "NAVY"


def try_query(label, params):
    print(f"\n--- {label} ---")
    print("params:", params)
    resp = requests.get(f"{BASE_URL}/items", headers=HEADERS, params=params)
    print("status:", resp.status_code)
    print("body:", resp.text[:1200])


# Attempt A: raw item -- no filter, just inspect one item's raw field names
try_query("A: one raw item (inspect field names / custom prefix)", {
    "~[primary_vendor_id]": VENDOR_ID,
    "per_page": 1,
})

# Attempt B: filter using custom@ (per docs)
try_query("B: custom@ prefix filter", {
    "~[primary_vendor_id]": VENDOR_ID,
    "~[custom@style_number]": STYLE_NUMBER,
    "~[custom@color]": COLOR,
    "per_page": 1,
})

# Attempt C: filter using item_custom@ prefix (in case it matches item_grids)
try_query("C: item_custom@ prefix filter", {
    "~[primary_vendor_id]": VENDOR_ID,
    "~[item_custom@style_number]": STYLE_NUMBER,
    "~[item_custom@color]": COLOR,
    "per_page": 1,
})
