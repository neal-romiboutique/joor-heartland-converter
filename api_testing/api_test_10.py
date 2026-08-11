#!/usr/bin/env python3
"""Find where a grid's dimension value lists live."""
import json, os, sys, requests
from pathlib import Path

env = Path(__file__).resolve().parent.parent / ".env"
for line in env.read_text().splitlines():
    if line.strip() and not line.startswith("#") and "=" in line:
        k, v = line.split("=", 1)
        os.environ.setdefault(k.strip(), v.strip())

BASE = f"https://{os.environ['HEARTLAND_SUBDOMAIN']}.retail.heartland.us/api"
H = {"Authorization": f"Bearer {os.environ['HEARTLAND_TOKEN']}"}
gid = sys.argv[1] if len(sys.argv) > 1 else "738"

print("=== A: single-resource fetch ===")
r = requests.get(f"{BASE}/item_grids/{gid}", headers=H, timeout=30)
print("HTTP", r.status_code)
if r.status_code == 200:
    d = r.json()
    d = d.get("result", d)
    print("keys:", sorted(d.keys()))
    for k in d:
        if "dim" in k.lower() or "cell" in k.lower():
            print(f"  {k}: {json.dumps(d[k])[:800]}")

print("\n=== B: items belonging to this grid ===")
for f in ("item_grid_id", "grid_id"):
    r = requests.get(f"{BASE}/items", headers=H,
                     params={f"~[{f}]": gid, "per_page": 3}, timeout=30)
    print(f"filter {f}: HTTP {r.status_code}", end=" ")
    if r.status_code == 200:
        d = r.json()
        print("total:", d.get("total"))
        if d.get("results"):
            print(json.dumps(d["results"][0], indent=2)[:2500])
            break
    else:
        print(r.text[:200])
