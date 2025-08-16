# ibkr_automation.py
# Read signals from result.xlsx, compute entries, and either print a plan (DryRun) or place orders via IBKR (ib_insync).
# IMPORTANT: Install dependencies on your machine: pip install ib-insync pandas openpyxl numpy
from dataclasses import dataclass
from typing import Optional, List, Tuple
import pandas as pd, numpy as np, time, math, pathlib, os

# ---- CONFIG ----
CONFIG_XLSX = "ibkr_config_template.xlsx"     # rename ibkr_config_template.xlsx to ibkr_config.xlsx and edit
RESULT_XLSX = "results.xlsx"          # your OCR output file
LEDGER_CSV  = "executed_signals.csv" # to dedupe signals between runs
PLAN_CSV    = "plan.csv"             # preview of orders
LOG_FILE    = "ibkr_automation.log"  # simple text log

# ---- Helpers ----
def log(msg: str):
    print(msg, flush=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(msg + "\n")

def ceil_to_tick(price: float, tick: float) -> float:
    return math.ceil(price / tick) * tick

def round_to_tick(price: float, tick: float) -> float:
    # Round to nearest tick
    return round(price / tick) * tick

def floor_to_tick(price: float, tick: float) -> float:
    return math.floor(price / tick) * tick

def parse_bool(s: str) -> bool:
    return str(s).strip().lower() in ("1","true","yes","y")

@dataclass
class RiskConfig:
    host: str = "127.0.0.1"
    port: int = 7497
    clientId: int = 123
    dry_run: bool = True
    use_paper: bool = True
    max_orders_per_signal: int = 15
    default_wait_sd: float = 0.0
    trailing_mode: str = "OFF"   # OFF or DYNAMIC_BPS
    one_percent_stop: bool = True

@dataclass
class SymbolConfig:
    symbol: str
    asset_class: str   # STK or FX
    exchange: str
    currency: str
    tick_size: float
    wait_sd: float
    entries_per_signal: int
    entry_mode: str    # LIMIT or TRAIL (TRAIL not implemented in v1)
    shares_per_entry: int
    fx_notional_per_entry: float
    enable: bool = True

def load_config(xlsx_path: str) -> Tuple[RiskConfig, pd.DataFrame]:
    xls = pd.ExcelFile(xlsx_path)
    sym = pd.read_excel(xls, "symbols").fillna(0)
    risk = pd.read_excel(xls, "risk")
    rdict = {row["Param"]: str(row["Value"]) for _,row in risk.iterrows()}
    rc = RiskConfig(
        host = rdict.get("Host","127.0.0.1"),
        port = int(rdict.get("Port","7497")),
        clientId = int(rdict.get("ClientId","123")),
        dry_run = parse_bool(rdict.get("DryRun","true")),
        use_paper = parse_bool(rdict.get("UsePaper","true")),
        max_orders_per_signal = int(rdict.get("MaxOrdersPerSignal", "15")),
        default_wait_sd = float(rdict.get("DefaultWaitSD","0.0")),
        trailing_mode = rdict.get("TrailingMode","OFF").upper(),
        one_percent_stop = parse_bool(rdict.get("OnePercentStop","true"))
    )
    # Normalize symbol config
    sym["Enable"] = sym["Enable"].apply(parse_bool)
    sym["WaitSD"] = sym.apply(lambda r: r["WaitSD"] if r["WaitSD"]!=0 else rc.default_wait_sd, axis=1)
    return rc, sym

def load_signals(xlsx_path: str) -> pd.DataFrame:
    df = pd.read_excel(xlsx_path)
    # Expect columns: SignalDate, SignalTime, Symbol, Signal, BidPrice, AskPrice, LastPrice, EqPrice, EqLevel, Bias
    needed = ["SignalDate","SignalTime","Symbol","Signal","BidPrice","AskPrice","LastPrice","EqPrice","EqLevel"]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns in {xlsx_path}: {missing}")
    # drop rows where bid/ask is zero
    df = df[(df["BidPrice"]>0) & (df["AskPrice"]>0)]
    # create a unique key per row for dedupe
    df["SignalKey"] = df["SignalDate"].astype(str)+" "+df["SignalTime"].astype(str)+"|"+df["Symbol"].astype(str)+"|"+df["Signal"].astype(str)
    return df

def load_ledger(path: str) -> set:
    if not os.path.exists(path):
        return set()
    ld = pd.read_csv(path)
    return set(ld["SignalKey"].tolist())

def append_ledger(path: str, keys: List[str]):
    mode = "a" if os.path.exists(path) else "w"
    pd.DataFrame({"SignalKey":keys}).to_csv(path, index=False, mode=mode, header=not os.path.exists(path))

def plan_orders(signals: pd.DataFrame, symcfg: pd.DataFrame, rc: RiskConfig) -> pd.DataFrame:
    plans = []
    symdict = {row["Symbol"]: row for _,row in symcfg.iterrows() if row["Enable"]}
    for _, s in signals.iterrows():
        if s["Symbol"] not in symdict:
            continue
        cfg = symdict[s["Symbol"]]
        wait_sd = float(cfg["WaitSD"])
        entries = int(min(cfg["EntriesPerSignal"], rc.max_orders_per_signal))
        sd_step = abs(float(s["EqLevel"]))  # price units per "std dev" step (assumption)
        if sd_step <= 0:
            continue
        tick = float(cfg["TickSize"])
        side = "BUY" if str(s["Signal"]).lower().startswith("long") else "SELL"
        base_price = float(s["BidPrice"]) if side=="BUY" else float(s["AskPrice"])

        for i in range(entries):
            k = wait_sd + i # number of SDs from base price
            if side=="BUY":
                limit_price = base_price - k*sd_step
                limit_price = round_to_tick(limit_price, tick)
                qty = int(cfg["SharesPerEntry"]) if cfg["AssetClass"]=="STK" else 0
                fx_notional = float(cfg["FXNotionalPerEntry"]) if cfg["AssetClass"]=="FX" else 0.0
            else:
                limit_price = base_price + k*sd_step
                limit_price = round_to_tick(limit_price, tick)
                qty = int(cfg["SharesPerEntry"]) if cfg["AssetClass"]=="STK" else 0
                fx_notional = float(cfg["FXNotionalPerEntry"]) if cfg["AssetClass"]=="FX" else 0.0

            plans.append({
                "SignalKey": s["SignalKey"],
                "When": f'{s["SignalDate"]} {s["SignalTime"]}',
                "Symbol": s["Symbol"],
                "AssetClass": cfg["AssetClass"],
                "Side": side,
                "BasePrice": base_price,
                "EqLevel": sd_step,
                "SDsFromBase": float(k),
                "PlannedPrice": limit_price,
                "TickSize": tick,
                "OrderType": cfg["EntryMode"],
                "Shares": qty,
                "FxNotional": fx_notional,
            })
    return pd.DataFrame(plans)

def main():
    rc, symcfg = load_config(CONFIG_XLSX)
    log(f"Loaded config from {CONFIG_XLSX}. DryRun={rc.dry_run}")
    sigs = load_signals(RESULT_XLSX)
    ledger = load_ledger(LEDGER_CSV)
    sigs = sigs[~sigs["SignalKey"].isin(ledger)]  # dedupe
    if sigs.empty:
        log("No new signals to process.")
        return

    plan = plan_orders(sigs, symcfg, rc)
    if plan.empty:
        log("No orders planned (check Enable, EqLevel, WaitSD).")
        return

    plan.to_csv(PLAN_CSV, index=False)
    log(f"Wrote plan to {PLAN_CSV} with {len(plan)} rows.")

    if rc.dry_run:
        log("DryRun=true, not sending orders.")
        # Mark signals as processed so they aren't re-planned; comment the next line if you prefer to re-plan until filled
        append_ledger(LEDGER_CSV, list(sigs["SignalKey"].unique()))
        return

    # --- Live placing via IBKR ---
    try:
        from ib_insync import IB, Stock, Forex, LimitOrder, MarketOrder, util, Contract
    except Exception as e:
        log("ib_insync not installed. Run: pip install ib-insync")
        raise

    ib = IB()
    ib.connect(rc.host, rc.port, clientId=rc.clientId)
    log("Connected to TWS/IBG.")

    # Build and place
    oids = []
    for _, row in plan.iterrows():
        symbol = row["Symbol"]
        if row["AssetClass"] == "STK":
            contract = Stock(symbol, exchange=symcfg.loc[symcfg["Symbol"]==symbol, "Exchange"].values[0],
                             currency=symcfg.loc[symcfg["Symbol"]==symbol, "Currency"].values[0])
            qty = int(row["Shares"])
        else:
            # FX symbols are like "EUR.USD"
            base, quote = symbol.split(".")
            contract = Forex(pair=f"{base}{quote}")  # ib_insync uses EURUSD
            qty = int(row["FxNotional"])  # notional in base currency units (IBKR expects size in base currency)
        side = row["Side"]
        limit_price = float(row["PlannedPrice"])
        if row["OrderType"] == "LIMIT":
            order = LimitOrder("BUY" if side=="BUY" else "SELL", qty, limit_price, tif="DAY")
        else:
            order = LimitOrder("BUY" if side=="BUY" else "SELL", qty, limit_price, tif="DAY")  # placeholder

        trade = ib.placeOrder(contract, order)
        oids.append(trade.order.orderId)
        log(f"Placed {row['OrderType']} {side} {qty} {symbol} @ {limit_price} (orderId {trade.order.orderId})")

    ib.disconnect()
    # Mark signals as processed
    append_ledger(LEDGER_CSV, list(sigs["SignalKey"].unique()))
    log("Done.")

if __name__ == "__main__":
    main()
