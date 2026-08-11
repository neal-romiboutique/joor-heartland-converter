#!/usr/bin/env python3
"""Test 3: the actual dedupe-check query. Given a vendor name, style number,
and color (the same fields we get from a JOOR order line), ask Heartland
directly whether a matching item grid already exists. Read-only."""
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

# --- use the known grid from test 2 as a positive-match test case ---
VENDOR_NAME = "Veronica Beard"
STYLE_NUMBER = "2109CD00210031"   # from test 2's output -> should be FOUND
COLOR = "NAVY"                    # from test 2's output -> should be FOUND

# a style/color that (almost certainly) does NOT exist, as a negative control
FAKE_STYLE_NUMBER = "ZZZZ-DOES-NOT-EXIST-99999"
FAKE_COLOR = "NOT-A-REAL-COLOR"


def resolve_vendor_id(name):
    resp = requests.get(
        f"{BASE_URL}/purchasing/vendors",
        headers=HEADERS,
        params={"query": name, "per_page": 5},
    )
    results = resp.json().get("results", [])
    if not results:
        raise SystemExit(f"No vendor found for '{name}'")
    return results[0]["id"], results[0]["name"]


def grid_exists(vendor_id, style_number, color):
    resp = requests.get(
        f"{BASE_URL}/item_grids",
        headers=HEADERS,
        params={
            "~[item_primary_vendor_id]": vendor_id,
            "~[custom@style_number]": style_number,
            "~[custom@color]": color,
            "per_page": 1,
        },
    )
    data = resp.json()
    total = data.get("total", 0)
    return resp.status_code, total, data


vendor_id, vendor_name = resolve_vendor_id(VENDOR_NAME)
print(f"Vendor: {vendor_name} (id={vendor_id})\n")

print("--- Positive control (should be FOUND) ---")
status, total, data = grid_exists(vendor_id, STYLE_NUMBER, COLOR)
print(f"style_number={STYLE_NUMBER!r} color={COLOR!r}")
print(f"status={status} total={total}")
if total:
    g = data["results"][0]
    print(f"  matched grid id={g['id']} description={g['description']!r}")

print("\n--- Negative control (should be NOT found) ---")
status, total, data = grid_exists(vendor_id, FAKE_STYLE_NUMBER, FAKE_COLOR)
print(f"style_number={FAKE_STYLE_NUMBER!r} color={FAKE_COLOR!r}")
print(f"status={status} total={total}")
