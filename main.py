import requests
import pandas as pd
import time
from datetime import datetime
from openpyxl import load_workbook

# API Endpoint for Top 50 Cryptocurrencies by Market Cap
URL = "https://api.coingecko.com/api/v3/coins/markets"
PARAMS = {
    "vs_currency": "usd",
    "order": "market_cap_desc",
    "per_page": 50,
    "page": 1,
    "sparkline": False,
    "price_change_percentage": "24h"
}

EXCEL_FILE = "Live_Crypto_Report.xlsx"


def fetch_crypto_data():
    """Fetches live cryptocurrency data from CoinGecko API."""
    response = requests.get(URL, params=PARAMS)
    if response.status_code == 200:
        data = response.json()
        return [
            {
                "Name": coin["name"],
                "Symbol": coin["symbol"].upper(),
                "Current Price (USD)": coin["current_price"],
                "Market Cap": coin["market_cap"],
                "24h Trading Volume": coin["total_volume"],
                "24h Price Change (%)": coin["price_change_percentage_24h"],
                "Last Updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            for coin in data
        ]
    else:
        print("Error fetching data:", response.status_code)
        return []


def update_excel():
    """Fetch and update Excel sheet with live data and analysis report."""
    crypto_data = fetch_crypto_data()
    if not crypto_data:
        return

    df = pd.DataFrame(crypto_data)

    # Perform basic analysis
    top_5 = df.nlargest(5, "Market Cap")
    avg_price = df["Current Price (USD)"].mean()
    highest_change = df.loc[df["24h Price Change (%)"].idxmax()]
    lowest_change = df.loc[df["24h Price Change (%)"].idxmin()]

    analysis_df = pd.DataFrame({
        "Metric": ["Top 5 Cryptos by Market Cap", "Average Price", "Highest 24h Change", "Lowest 24h Change", "Last Updated"],
        "Value": [top_5["Name"].tolist(), avg_price, highest_change["Name"], lowest_change["Name"], datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
    })

    try:
        # Load existing file and update
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, index=False, sheet_name="Live Data")
            analysis_df.to_excel(writer, index=False, sheet_name="Analysis Report")
    except FileNotFoundError:
        # Create new file if not exists
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Live Data")
            analysis_df.to_excel(writer, index=False, sheet_name="Analysis Report")

    print(f"Excel and Report updated successfully at {datetime.now().strftime('%H:%M')}!")


def run_continuously(interval=300):
    """Runs the script every 'interval' seconds (default: 5 minutes)."""
    while True:
        update_excel()
        print(f"Next run at {datetime.now().strftime('%H:%M')} + {interval // 60} minutes")
        time.sleep(interval)


if __name__ == "__main__":
    run_continuously()