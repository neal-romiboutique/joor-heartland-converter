#!/usr/bin/env python3
"""Test 5: item_grids' custom fields appear as "item_custom" in the response
body (not "custom"), so try filtering with the "item_custom@" prefix instead
of "custom@". Read-only."""
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
STYLE_NUMBER = "2109CD00210031"
COLOR = "NAVY"


def try_query(label, params):
    print(f"\n--- {label} ---")
    print("params:", params)
    resp = requests.get(f"{BASE_URL}/item_grids", headers=HEADERS, params=params)
    print("status:", resp.status_code)
    print("body:", resp.text[:800])


# Attempt D: item_custom@ prefix, positive control (should be FOUND)
try_query("D: item_custom@ prefix (positive control)", {
    "~[item_primary_vendor_id]": VENDOR_ID,
    "~[item_custom@style_number]": STYLE_NUMBER,
    "~[item_custom@color]": COLOR,
    "per_page": 1,
})

# Attempt E: same prefix, negative control (should be NOT found)
try_query("E: item_custom@ prefix (negative control)", {
    "~[item_primary_vendor_id]": VENDOR_ID,
    "~[item_custom@style_number]": "ZZZZ-DOES-NOT-EXIST-99999",
    "~[item_custom@color]": "NOT-A-REAL-COLOR",
    "per_page": 1,
})
