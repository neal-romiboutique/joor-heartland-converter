#!/usr/bin/env python3
"""Dump one item_grid's full JSON so we can map API fields -> import columns."""
import json, os, sys, requests
from pathlib import Path

env = Path(__file__).resolve().parent.parent / ".env"
for line in env.read_text().splitlines():
    if line.strip() and not line.startswith("#") and "=" in line:
        k, v = line.split("=", 1)
        os.environ.setdefault(k.strip(), v.strip())

BASE = f"https://{os.environ['HEARTLAND_SUBDOMAIN']}.retail.heartland.us/api"
H = {"Authorization": f"Bearer {os.environ['HEARTLAND_TOKEN']}"}

grid_no = sys.argv[1] if len(sys.argv) > 1 else "738"
r = requests.get(f"{BASE}/item_grids", headers=H,
                 params={"~[public_id]": grid_no, "per_page": 1}, timeout=30)
print("HTTP", r.status_code)
d = r.json()
print("total:", d.get("total"))
res = d.get("results", [])
if res:
    print(json.dumps(res[0], indent=2)[:6000])
else:
    print("No match. Try: ~[id] instead of ~[public_id], or drop the filter")
    print("and use per_page=1 to see any grid's shape.")

