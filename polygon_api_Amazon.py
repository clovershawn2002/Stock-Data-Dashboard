import requests
import pandas as pd
import pprint
import os

# directory where all output Excel files will be stored (OneDrive StockData)
onedrive_dir = os.path.join(os.environ.get("USERPROFILE", r"C:\Users\YourName"), "OneDrive", "StockData")
os.makedirs(onedrive_dir, exist_ok=True)

BASE_URL = "https://api.massive.com"  # No trailing slash
API_KEY = "AIBKhvM0NQ4so4CQojCu9JGQaqx_1nQj"
TICKER = "AMZN"

# Mapping of tickers to company names
TICKER_NAMES = {
    "GOOGL": "Google Inc.",
    "AAPL": "Apple Inc.",
    "MSFT": "Microsoft Corporation",
    "AMZN": "Amazon.com Inc.",
    "TSLA": "Tesla Inc."
}

STOCK_NAME = TICKER_NAMES.get(TICKER, TICKER)  # Default to ticker if not in mapping

# 1. Ticker Overview
Ticker_Overview_Endpoint = f"{BASE_URL}/v3/reference/tickers/{TICKER}"
Ticker_Overview_Params = {"apiKey": "AIBKhvM0NQ4so4CQojCu9JGQaqx_1nQj"}

response = requests.get(Ticker_Overview_Endpoint, params=Ticker_Overview_Params, timeout=10)
Ticker_Overview_data = response.json()

df_Ticker_Overview_data = pd.DataFrame([Ticker_Overview_data["results"]])


# 2. Custom bars
from datetime import datetime

from_date = "2025-05-01"  # you choose your start date
to_date = datetime.today().strftime("%Y-%m-%d") 
Custom_bars_Endpoint = f"{BASE_URL}/v2/aggs/ticker/{TICKER}/range/1/day/{from_date}/{to_date}"
Custom_bars_Params = {"apiKey": "AIBKhvM0NQ4so4CQojCu9JGQaqx_1nQj",
"adjusted": "true", 
"limit": 1200}

response = requests.get(Custom_bars_Endpoint, params=Custom_bars_Params)
Custom_bars_data = response.json()
Custom_bars = Custom_bars_data["results"]
df_Custom_bars_data = pd.DataFrame(Custom_bars)
# Polygon returns epoch timestamps (milliseconds) in the 't' field for bars.
# Convert to readable Date and Time columns (separate) if present.
if "t" in df_Custom_bars_data.columns:
    # Convert milliseconds since epoch to pandas datetime
    dt = pd.to_datetime(df_Custom_bars_data["t"], unit='ms')
    # Create separate string columns for Date and Time (safe for Excel export)
    df_Custom_bars_data["Date"] = dt.dt.strftime("%Y-%m-%d")
    df_Custom_bars_data["Time"] = dt.dt.strftime("%H:%M:%S")
    # Move Date and Time to the front
    cols = ["Date", "Time"] + [c for c in df_Custom_bars_data.columns if c not in ("Date", "Time")]
    df_Custom_bars_data = df_Custom_bars_data[cols]

# Rename OHLC columns if present and keep a sensible column order
rename_map = {"o": "Open", "h": "High", "l": "Low", "c": "Close"}
existing_rename = {k: v for k, v in rename_map.items() if k in df_Custom_bars_data.columns}
if existing_rename:
    df_Custom_bars_data.rename(columns=existing_rename, inplace=True)

# Reorder columns to Date, Time, Open, High, Low, Close if they exist
preferred_order = [col for col in ["Date", "Time", "Open", "High", "Low", "Close"] if col in df_Custom_bars_data.columns]
other_cols = [c for c in df_Custom_bars_data.columns if c not in preferred_order]
df_Custom_bars_data = df_Custom_bars_data[preferred_order + other_cols]


# 3. Pervious_Day_Bar

# Endpoint
Pervious_Day_Bar_Endpoint = f"{BASE_URL}/v2/aggs/ticker/{TICKER}/prev"
Pervious_Day_Bar_Params = {"apiKey": API_KEY}

# Request
response = requests.get(Pervious_Day_Bar_Endpoint, params=Pervious_Day_Bar_Params)
Pervious_Day_Bar_data = response.json()

# Print full JSON for checking
pprint.pprint(Pervious_Day_Bar_data)

# Extract results safely
if "results" in Pervious_Day_Bar_data and len(Pervious_Day_Bar_data["results"]) > 0:
    close_price = Pervious_Day_Bar_data["results"][0]["c"]

    # Build DataFrame with Ticker + Close Price
    df = pd.DataFrame([{
        "Close_Price": close_price
    }])

    # Save to Excel
    
    #print(df)
else:
    print(" No results found in API response")




# 4. MACD
MACD_Endpoint = f"{BASE_URL}/v1/indicators/macd/{TICKER}"
# Initialize as empty DataFrame (will be populated if API succeeds)
df_MACD = pd.DataFrame()
# Make MACD timeframe match Custom Bars (daily) and request enough periods to cover the same range
from datetime import datetime as _dt
days_span = (_dt.today() - _dt.strptime(from_date, "%Y-%m-%d")).days + 10
MACD_Params = { 
    "apiKey": API_KEY,
    "timespan": "day",            # match custom bars daily granularity
    "adjusted": "true",
    "short_window": "15",
    "long_window": "30",
    "signal_window": "7",
    "series_type": "close",
    "limit": str(max(days_span, 100))  # request at least days_span (string accepted by API)
}

response = requests.get(MACD_Endpoint, params=MACD_Params)
MACD_data = response.json()

# 3️⃣ Check structure
# pprint.pprint(MACD_data)

# 4️⃣ Extract list of MACD values
if "results" in MACD_data and "values" in MACD_data["results"]:
    macd_values = MACD_data["results"]["values"]

    # Create DataFrame
    df_MACD = pd.DataFrame(macd_values)

    # Convert timestamp to readable Date (match Custom_bars Date column)
    if "timestamp" in df_MACD.columns:
        df_MACD["Date"] = pd.to_datetime(df_MACD["timestamp"], unit='ms').dt.strftime("%Y-%m-%d")
        # move Date to front
        cols = ["Date"] + [c for c in df_MACD.columns if c != "Date"]
        df_MACD = df_MACD[cols]

    # Rename columns for clarity
    df_MACD.rename(columns={
        "timestamp": "Timestamp",
        "macd": "MACD",
        "signal": "Signal",
        "histogram": "Histogram"
    }, inplace=True)

    # 5️⃣ Save to Excel
    #df_MACD.to_excel("MACD_Data.xlsx", index=False)
   # print(df_MACD)
else:
    print(" No MACD data found in response.")



# 6. Financial
Financial_Endpoint = f"{BASE_URL}/vX/reference/financials"
Financial_Params = {"apiKey": API_KEY,
"ticker": TICKER,
"timeframe":"annual",
"limit":5,
"order":"desc"
}

response = requests.get(Financial_Endpoint, params=Financial_Params)
Financial_data = response.json()

pprint.pprint(Financial_data)
print(f"Request URL: {response.url}")
if "results" not in Financial_data or len (Financial_data["results"]) == 0:
    print(" No financial data found in response.")
    print(f"Response status: {response.status_code}")
    print(f"Full response: {Financial_data}")
    exit()
else:
    result = Financial_data["results"][0]["financials"]

    revenue = result["income_statement"]["revenues"]["value"]
    gross_profit = result["income_statement"]["gross_profit"]["value"]
    Total_Asset = result["balance_sheet"]["assets"]["value"]
    Operating_Cash_Flow = result["cash_flow_statement"]["net_cash_flow_from_operating_activities"]["value"]
    Financial_data_row = {"revenue": revenue,
    "gross_profit": gross_profit,
    "Total_Asset": Total_Asset,
    "Operating_Cash_Flow": Operating_Cash_Flow}
    df_Financial_data_row = pd.DataFrame([Financial_data_row])



# 7. Ticker Overview (Market Cap and OutstandingShares)
Ticker_Overview_Endpoint = f"{BASE_URL}/v3/reference/tickers/{TICKER}"
Ticker_Overview_Params = {"apiKey": "AIBKhvM0NQ4so4CQojCu9JGQaqx_1nQj"}

response = requests.get(Ticker_Overview_Endpoint, params=Ticker_Overview_Params)
Ticker_Overview_data = response.json()

#pprint.pprint(Ticker_Overview_data)
#result = Financial_data["results"][0]["market_cap"]
if "results" in Ticker_Overview_data:
    Market_Cap = Ticker_Overview_data["results"]["market_cap"]
    Outstanding_Shares = Ticker_Overview_data["results"]["share_class_shares_outstanding"]
    Ticker_Overview_data_row = {"market_cap": Market_Cap,
    "Outstanding_Shares": Outstanding_Shares}
    df_Ticker_Overview_data_row = pd.DataFrame([Ticker_Overview_data_row])
else:
    print(" No Ticker Overview data found in response.")
    exit()




# Use dynamic sheet names with TICKER prefix (max 31 chars)
sheet_prefix = TICKER  # "GOOGL" or "AAPL"
# build path inside OneDrive folder
output_path = os.path.join(onedrive_dir, f"{sheet_prefix}.xlsx")

# Add "Stock Name" column to all DataFrames (insert at position 0)
df_Ticker_Overview_data.insert(0, "Stock Name", STOCK_NAME)
df_Custom_bars_data.insert(0, "Stock Name", STOCK_NAME)
df.insert(0, "Stock Name", STOCK_NAME)
df_MACD.insert(0, "Stock Name", STOCK_NAME)
df_Financial_data_row.insert(0, "Stock Name", STOCK_NAME)
df_Ticker_Overview_data_row.insert(0, "Stock Name", STOCK_NAME)

with pd.ExcelWriter(output_path) as writer:
    df_Ticker_Overview_data.to_excel(writer, sheet_name=f"{sheet_prefix}_Ticker", index=False)
    # Transpose without Stock Name, then add it back at the end
    stock_name_value = df_Ticker_Overview_data["Stock Name"].iloc[0]
    df_transposed = df_Ticker_Overview_data.drop(columns=["Stock Name"]).T
    df_transposed["Stock Name"] = stock_name_value
    df_transposed.to_excel(writer, sheet_name=f"{sheet_prefix}_Ticker_T", index=True)
    df_Custom_bars_data.to_excel(writer, sheet_name=f"{sheet_prefix}_Bars", index=False)
    df.to_excel(writer, sheet_name=f"{sheet_prefix}_PrevDay", index=False)
    df_MACD.to_excel(writer, sheet_name=f"{sheet_prefix}_MACD", index=False)
    df_Financial_data_row.to_excel(writer, sheet_name=f"{sheet_prefix}_Fin", index=False)
    df_Ticker_Overview_data_row.to_excel(writer, sheet_name=f"{sheet_prefix}_OV_Data", index=False)

print(f" Data saved to {output_path}")

