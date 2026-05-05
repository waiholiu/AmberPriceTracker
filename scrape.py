"""
Fetches live energy prices for postcode 2600 from the Amber public backend
(https://backend.amber.com.au/postcode/{postcode}/prices) and appends a
timestamped row to a local Excel file (committed back to the repo).
"""

import json
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import Workbook, load_workbook

POSTCODE = "2600"
API_URL = f"https://backend.amber.com.au/postcode/{POSTCODE}/prices?past-hours=1"
XLSX_PATH = Path("Data/prices.xlsx")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/136.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json",
    "Origin": "https://www.amber.com.au",
    "Referer": "https://www.amber.com.au/",
}


def _latest_interval(price_blocks: list) -> dict | None:
    """Pick the interval with the most recent nemTime across all blocks."""
    latest = None
    for block in price_blocks or []:
        for interval in block.get("intervals", []):
            if latest is None or interval.get("nemTime", "") > latest.get("nemTime", ""):
                latest = interval
    return latest


def fetch_prices() -> dict:
    print(f"  Fetching {API_URL} ...")
    req = urllib.request.Request(API_URL, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=30) as resp:
        payload = json.load(resp)

    energy_interval = _latest_interval(payload.get("priceData", []))
    feedin_interval = _latest_interval(payload.get("feedInPriceData", []))

    return {
        "energy_price_c_kwh": energy_interval.get("perKwh") if energy_interval else None,
        "feed_in_rate_c_kwh": feedin_interval.get("perKwh") if feedin_interval else None,
        "nem_time": (energy_interval or feedin_interval or {}).get("nemTime"),
    }


def main():
    print(f"Fetching Amber energy prices for postcode {POSTCODE} ...")
    data = fetch_prices()

    if data["energy_price_c_kwh"] is None and data["feed_in_rate_c_kwh"] is None:
        raise RuntimeError("Amber backend returned no price intervals")

    print(
        f"  Live energy price : {data['energy_price_c_kwh']} c/kWh\n"
        f"  Live feed-in rate : {data['feed_in_rate_c_kwh']} c/kWh\n"
        f"  NEM time          : {data['nem_time']}"
    )

    timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

    XLSX_PATH.parent.mkdir(parents=True, exist_ok=True)

    if XLSX_PATH.exists():
        wb = load_workbook(XLSX_PATH)
        ws = wb.active
    else:
        print("  No existing file — creating new spreadsheet.")
        wb = Workbook()
        ws = wb.active
        ws.title = "Amber Prices"
        ws.append(["Timestamp", "Energy Price (c/kWh)", "Feed-In Rate (c/kWh)"])

    ws.append([timestamp, data["energy_price_c_kwh"], data["feed_in_rate_c_kwh"]])

    wb.save(XLSX_PATH)
    print(f"Done! Row added at {timestamp} → {XLSX_PATH}")


if __name__ == "__main__":
    main()
