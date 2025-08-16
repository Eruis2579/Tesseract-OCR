from ib_insync import Stock,Forex,LimitOrder,StopOrder,Order

def place_limit_order(ib, symbol, asset_type, qty, price, action):
    """
    asset_type: 'Stock' or 'Forex'
    action: 'BUY' or 'SELL'
    """
    if asset_type == "Stock":
        contract = Stock(symbol, 'SMART', 'USD')
    elif asset_type == "Forex":
        base, quote = symbol[:3], symbol[-3:]
        print(f"{base}{quote}")
        contract = Forex(f"{base}{quote}")
    else:
        print(f"❌ Unknown asset type for {symbol}")
        return None

    order = LimitOrder(action, qty, price)
    trade = ib.placeOrder(contract, order)
    print(f"📤 Placed {action} {qty} {symbol} @ {price}")
    return trade

def place_stop_loss(ib, symbol, asset_type, qty, stop_price, action):
    if asset_type == "Stock":
        contract = Stock(symbol, 'SMART', 'USD')
    elif asset_type == "Forex":
        base, quote = symbol[:3], symbol[-3:]
        contract = Forex(f"{base}{quote}")
    else:
        return None

    sl_action = 'SELL' if action == 'BUY' else 'BUY'
    order = StopOrder(sl_action, qty, stop_price)
    trade = ib.placeOrder(contract, order)
    print(f"📉 Stop Loss set for {symbol} @ {stop_price}")
    return trade
    
def place_trailing_stop(ib, symbol, asset_type, qty, trail_amount, action):
    if asset_type == "Stock":
        contract = Stock(symbol, 'SMART', 'USD')
    elif asset_type == "Forex":
        base, quote = symbol[:3], symbol[-3:]
        contract = Forex(f"{base}{quote}")
    else:
        return None

    sl_action = 'SELL' if action == 'BUY' else 'BUY'
    order = Order()
    order.action = sl_action
    order.totalQuantity = qty
    order.orderType = 'TRAIL'
    order.trailingAmount = trail_amount
    order.tif = 'GTC'  # Good till canceled
    trade = ib.placeOrder(contract, order)
    print(f"📉 Trailing Stop set for {symbol}, trail {trail_amount}")
    return trade

def cancel_all_orders_for_symbol(ib, symbol):
    open_trades = ib.trades()
    for t in open_trades:
        if t.contract.symbol == symbol:
            ib.cancelOrder(t.order)
            print(f"🛑 Canceled order {t.order.orderId} for {symbol}")