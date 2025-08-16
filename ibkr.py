import pandas as pd
from config import *
import math
from datetime import datetime
from order import place_limit_order, cancel_all_orders_for_symbol, place_trailing_stop
from ib_insync import *
import logging
from pathlib import Path
import time

xlsx_path = XLSX_PATH
tick_path = TICK_PATH

logging.basicConfig(
    filename="trading.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

def connect_ibkr():
    ib = IB()
    ib.connect(IB_HOST, IB_PORT, clientId=IB_CLIENT_ID)
    print("✅ Connected to IBKR")
    return ib

def clean_signals() -> pd.DataFrame:
    df = pd.read_excel(xlsx_path)
    # Step 2a: Remove duplicates (SignalDate + SignalTime + Symbol as unique key)
    df["UniqueKey"] = df["SignalDate"].astype(str) + "_" + df["SignalTime"].astype(str) + "_" + df["Symbol"]
    df = df.drop_duplicates(subset=["UniqueKey"], keep="last")

    # Step 2b: Remove rows with BidPrice or AskPrice = 0
    df = df[(df["BidPrice"] != 0) & (df["AskPrice"] != 0)]
    # Step 2c: Convert numeric columns safely
    numeric_cols = ["BidPrice", "AskPrice", "LastPrice", "EqPrice", "EqLevel", "Bias"]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Drop rows where any critical numeric value is missing
    df = df.dropna(subset=numeric_cols)

    # Drop helper column
    df = df.drop(columns=["UniqueKey"])
    return df

def load_symbol_config():
    """
    Expected Excel sheet format:
    Symbol | AssetType | WaitDevs | MaxOrders | PercentCapital | FixedShares | FixedForexUSD
    AAPL   | equity    | 1        | 15        | 0.02           | 50          | 
    EURUSD | forex     | 1        | 10        | 0.02           |             | 100000
    """
    df = pd.read_excel(TICK_PATH, sheet_name="Sheet1")
    df = df.set_index("Symbol").to_dict(orient="index")
    return df

def calculate_position_size(symbol, config, portfolio_value, leverage):
    settings = config.get(symbol, {})
    asset_type = settings.get("Type", "Stock")

    if asset_type == "Stock":
        # Portfolio × leverage × 80% × PercentCapital
        percent_capital = settings.get("PercentCapital", 0.02)  # default 2%
        capital_available = portfolio_value * leverage * 0.80
        entry_value = capital_available * percent_capital
        return math.floor(entry_value / settings.get("LastPrice", 1))

    elif asset_type == "Forex":
        # For testing, use fixed USD amount (e.g., $100k)
        return settings.get("FixedForexUSD", 100000)

    return 0

def get_entry_price(signal, bid, ask):
    if signal == "LongTrigger":
        return bid
    elif signal == "ShortTrigger":
        return ask
    return None

def compute_sd_tick(eq_price, last_price, eq_level, tick_size):
    if eq_level == 0:
        return 0
    raw_sd = abs(eq_price - last_price) / abs(eq_level)
    return int(math.ceil(raw_sd / tick_size)) # round up to nearest tick

def compute_sd(eq_price, last_price, eq_level, tick_size):
    if eq_level == 0:
        return 0
    raw_sd = abs(eq_price - last_price) / abs(eq_level)
    return math.ceil(raw_sd / tick_size) * tick_size # round up to nearest tick

def generate_pyramid_orders(row):
    """Generate list of prices for pyramid entries."""
    orders = []
    if pd.isna(row["EntryPrice"]) or row["SD"] <= 0:
        return orders

    for i in range(row["MaxOrders"]):
        deviation = (row["WaitDevs"] + i) * row["SD"]
        if row["Signal"] == "LongTrigger":
            price = row["EntryPrice"] + deviation
        else:  # ShortTrigger
            price = row["EntryPrice"] - deviation
        orders.append(round(price, 5))  # round to tick precision
    return orders

def stop_loss_progression(profit_bps):
    """
    Dynamic trailing stop logic:
    - 4 bps → SL = 2 bps
    - 6 bps → SL = 4 bps
    - 10 bps → SL = 6 bps
    - 15 bps → SL = 10 bps
    - 25 bps → SL = 15 bps
    - Continue pattern: +10 bps profit → +10 bps higher start, SL = prev_profit - 10
    """
    levels = [(4, 2), (6, 4), (10, 6), (15, 10), (25, 15)]
    sl=2
    if profit_bps < 35:
        for p, s in reversed(levels):
            if profit_bps >= p:
                sl=s
        return sl
    else:
        last_sl = 35+10*(int((profit_bps-35)/10))
        return last_sl-10

def calc_stop_loss(row):
    # Calculate bps change
    if row["Signal"] == "LongTrigger":
        profit_bps = ((row["LastPrice"] - row["EntryPrice"]) / row["EntryPrice"]) * 10000
    else:
        profit_bps = ((row["EntryPrice"] - row["LastPrice"]) / row["EntryPrice"]) * 10000


    # 1% account value hard stop
    one_percent_stop = -1  # in % of account value (bps ≈ 100bps per 1%)
    if profit_bps <= one_percent_stop * 100:
        return "EXIT_1PCT"
    # if(row['StopLossAction'] is not None and row['StopLossAction'] != ""):
    #     if(profit_bps<=int(row['StopLossAction'])):
    #         return "EXIT_2PCT"
    
    sl_bps = stop_loss_progression(profit_bps)
    if sl_bps is not None:
        return f"TRAIL_SL_{sl_bps}"

def result_with_sd(df: pd.DataFrame) -> pd.DataFrame:
    tick_sizes = pd.read_excel(tick_path, sheet_name="Sheet1").set_index("Symbol")["QuoteTick"].to_dict()
    df["TickSize"] = df["Symbol"].map(tick_sizes).fillna(0.0001)
    df["SD"] = df.apply(lambda row: compute_sd(row["EqPrice"], row["LastPrice"], row["EqLevel"], row["TickSize"]), axis=1)
    df["SD_tick"] = df.apply(lambda row: compute_sd_tick(row["EqPrice"], row["LastPrice"], row["EqLevel"], row["TickSize"]), axis=1)
    df = df[df["SD"] > 0]

    # ---- Step 4: Trade validation ----
    # Custom wait deviations (for now fixed, later from table)
    symbol_config = load_symbol_config()
    df["WaitDevs"] = df["Symbol"].apply(lambda s: symbol_config.get(s, {}).get("WaitDevs", 1))
    df["MaxOrders"] = df["Symbol"].apply(lambda s: symbol_config.get(s, {}).get("MaxOrders", 5))
    df["PositionSize"] = df.apply(
        lambda row: calculate_position_size(
            row["Symbol"], symbol_config,
            portfolio_value=100000,  # testing value
            leverage=3 if symbol_config.get(row["Symbol"], {}).get("Type") == "Stock" else 30
        ),
        axis=1
    )

    df["EntryPrice"] = df.apply(lambda row: get_entry_price(row["Signal"], row["BidPrice"], row["AskPrice"]), axis=1)
    df["LastUpdated"] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    df["PyramidOrders"] = df.apply(lambda row: generate_pyramid_orders(row), axis=1)

    df["StopLossAction"] = df.apply(lambda row: calc_stop_loss(row), axis=1)

    # Mark cancel conditions (exit reached)
    df["CancelRemainingOrders"] = df["StopLossAction"].apply(lambda x: True if x and x.startswith("EXIT") else False)
    return df

def main():
    ib = connect_ibkr()
    df = clean_signals()
    sd_df = result_with_sd(df)
    symbol_config = load_symbol_config()
    sd_df.to_excel(SD_CLEANED_PATH, index=False)
    for _, row in sd_df.iterrows():
        symbol = row["Symbol"]
        settings = symbol_config.get(symbol, {})
        asset_type = settings.get("Type", "Stock")
        qty = row["PositionSize"]

        # Place pyramid limit orders
        for price in row["PyramidOrders"]:
            action = "BUY" if row["Signal"] == "LongTrigger" else "SELL"
            place_limit_order(ib, symbol, asset_type, qty, price, action)

        # Stop loss handling
        if row["StopLossAction"] and row["StopLossAction"].startswith("TRAIL_SL"):
            trail_amount = float(row["StopLossAction"].replace("TRAIL_SL_", "").replace("bps", "")) / 10000
            place_trailing_stop(ib, symbol, asset_type, qty, trail_amount, action)
        elif row["StopLossAction"] == "EXIT_1PCT":
            cancel_all_orders_for_symbol(ib, symbol)

    ib.sleep(2)  # allow orders to be sent
    ib.disconnect()
    print(sd_df)

def process_signals(ib, portfolio_value=100000):
    """
    Core signal processing for one iteration.
    Loads signals, calculates SD, places/cancels orders.
    """
    try:
        if not Path(xlsx_path).exists():
            logging.warning("⚠️ Signal file not found")
            return

        df = clean_signals()
        if df.empty:
            logging.info("No signals found in Excel")
            return

        sd_df = result_with_sd(df)
        symbol_config = load_symbol_config()
        sd_df.to_excel(SD_CLEANED_PATH, index=False)

        for _, row in sd_df.iterrows():
            symbol = row["Symbol"]
            settings = symbol_config.get(symbol, {})
            asset_type = settings.get("Type", "Stock")
            qty = row["PositionSize"]

            # --- Place pyramid limit orders ---
            for price in row["PyramidOrders"]:
                action = "BUY" if row["Signal"] == "LongTrigger" else "SELL"
                place_limit_order(ib, symbol, asset_type, qty, price, action)

            # --- Stop loss handling ---
            if row["StopLossAction"] and row["StopLossAction"].startswith("TRAIL_SL"):
                trail_amount = float(row["StopLossAction"].replace("TRAIL_SL_", "").replace("bps", "")) / 10000
                place_trailing_stop(ib, symbol, asset_type, qty, trail_amount, action)
            elif row["StopLossAction"] == "EXIT_1PCT":
                cancel_all_orders_for_symbol(ib, symbol)

        logging.info(f"✅ Processed {len(sd_df)} signals")

    except Exception as e:
        logging.error(f"❌ Error in process_signals: {e}", exc_info=True)

def automation_loop():
    ib = connect_ibkr()

    try:
        while True:
            process_signals(ib, portfolio_value=100000)
            time.sleep(60)  # check every 60 seconds
    except KeyboardInterrupt:
        print("🛑 Automation stopped manually")
    finally:
        ib.disconnect()
        print("🔌 Disconnected from IBKR")

if __name__ == "__main__":
    automation_loop()