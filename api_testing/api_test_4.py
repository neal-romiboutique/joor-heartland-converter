
#!/usr/bin/env python3
"""Test 4: debug the 400 error from test 3 by printing Heartland's raw error
body, and try a couple of alternate filter syntaxes. Read-only."""
import os
import json
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

VENDOR_ID = 100387  # Veronica Beard, from test 2/3
STYLE_NUMBER = "2109CD00210031"
COLOR = "NAVY"


def try_query(label, params):
    print(f"\n--- {label} ---")
    print("params:", params)
    resp = requests.get(f"{BASE_URL}/item_grids", headers=HEADERS, params=params)
    print("status:", resp.status_code)
    print("body:", resp.text[:1000])


# Attempt A: original bracket-style params (what failed in test 3)
try_query("A: ~[field] bracket params", {
    "~[item_primary_vendor_id]": VENDOR_ID,
    "~[custom@style_number]": STYLE_NUMBER,
    "~[custom@color]": COLOR,
    "per_page": 1,
})

# Attempt B: JSON-encoded _filter param instead
filt = {
    "item_primary_vendor_id": VENDOR_ID,
    "custom@style_number": STYLE_NUMBER,
    "custom@color": COLOR,
}
try_query("B: JSON _filter param", {
    "_filter": json.dumps(filt),
    "per_page": 1,
})

# Attempt C: vendor filter only (isolate whether custom@ is the problem)
try_query("C: vendor id only, no custom fields", {
    "~[item_primary_vendor_id]": VENDOR_ID,
    "per_page": 1,
})

