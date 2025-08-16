from ib_insync import *
import pandas as pd
import time
from datetime import datetime
from config import *
# ========= SETTINGS =========
EXCEL_FILE = TICK_PATH
UPDATE_INTERVAL = TICK_UPDATE_INTERVAL  # seconds

# ========= FUNCTIONS =========
def get_order_tick_size(contract):
    details = ib.reqContractDetails(contract)
    if not details:
        return None
    marketRuleId = details[0].marketRuleIds.split(',')[0]
    if not marketRuleId:
        return None
    ticks = ib.reqMarketRule(int(marketRuleId))
    return ticks[0].increment if ticks else None

def get_quote_tick_size(symbol, type_):
    if type_.lower() != 'forex':
        return None  # For stocks, quote tick = order tick
    # Forex: standard market precision
    if symbol.endswith('JPY'):
        return 0.01
    else:
        return 0.0001

def update_tick_sizes():
    df = pd.read_excel(EXCEL_FILE)

    for i, row in df.iterrows():
        symbol = row['RealSymbol']
        type_ = row['Type']
        
        if type_.lower() == 'stock':
            contract = Stock(symbol, 'SMART', 'USD')
        elif type_.lower() == 'forex':
            contract = Forex(symbol)
        else:
            continue

        order_tick = get_order_tick_size(contract)
        quote_tick = get_quote_tick_size(symbol, type_)

        df.at[i, 'OrderTick'] = order_tick
        df.at[i, 'QuoteTick'] = quote_tick if quote_tick is not None else order_tick
        df.at[i, 'LastUpdated'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    df.to_excel(EXCEL_FILE, index=False)
    print(f"[{datetime.now()}] Updated tick sizes in {EXCEL_FILE}")

# ========= MAIN LOOP =========
ib = IB()
ib.connect(IB_HOST, IB_PORT, clientId=IB_CLIENT_ID)  # 7497 = paper

while True:
    try:
        update_tick_sizes()
        time.sleep(UPDATE_INTERVAL)
    except KeyboardInterrupt:
        print("Stopping...")
        break

ib.disconnect()
