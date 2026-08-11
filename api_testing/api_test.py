#!/usr/bin/env python3
"""Quick connectivity test for the Heartland Retail API. Read-only."""
import os
import requests
from pathlib import Path

# --- load .env manually (no extra dependency) ---
env_path = Path(__file__).resolve().parent / ".env"
for line in env_path.read_text().splitlines():
    if line.strip() and not line.startswith("#") and "=" in line:
        k, v = line.split("=", 1)
        os.environ.setdefault(k.strip(), v.strip())

TOKEN = os.environ["HEARTLAND_TOKEN"]
SUBDOMAIN = os.environ["HEARTLAND_SUBDOMAIN"]
BASE_URL = f"https://{SUBDOMAIN}.retail.heartland.us/api"

resp = requests.get(
    f"{BASE_URL}/purchasing/vendors",
    headers={"Authorization": f"Bearer {TOKEN}"},
    params={"per_page": 1},
)
print("Status:", resp.status_code)
print(resp.text[:500])
