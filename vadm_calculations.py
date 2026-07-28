"""
VADM Calculation Layer
======================
Pure calculation functions - no Streamlit/UI code here.
Every function takes parameters explicitly; nothing is hardcoded that should
be a UI toggle.

STATUS OF EACH SECTION:
  - EOD2 data loading........... TESTED (verified against live raw.githubusercontent.com)
  - Delivery % / relative........TESTED logic, verified against real EOD2 columns
  - Screener Excel parsing.......TESTED against your actual uploaded IRB_Infra_Devl-3.xlsx
  - EV / Total Debt / PE-per-year TESTED against IRB data (formulas only - sanity-check
                                  one year's PE against Screener.in's own chart)
  - Promoter holding (nse pkg)...NOT TESTED. Sandbox network blocks nseindia.com,
                                  so this is written correctly per the documented
                                  nse.shareholding() signature but never actually
                                  executed. Run it yourself before trusting it.
  - Alpha White sector-avg PE....INCOMPLETE ON PURPOSE. The comparison logic is
                                  here; where sector_avg_pe itself comes from is
                                  still an open decision (see bottom of file).
  - Alpha Black / VADM_t.........NOT BUILT. Formula f() and H4 were never given -
                                  this is flagged in your own project memory as a
                                  CEO-level decision. Nothing here should be
                                  invented to fill that gap.
"""

import json
import pandas as pd
import numpy as np
import openpyxl


# ---------------------------------------------------------------------------
# 1. STOCK UNIVERSE
# ---------------------------------------------------------------------------

def fetch_stock_universe() -> dict:
    """
    Returns {symbol: isin} for ~4500 NSE-listed symbols.
    Source verified live: raw.githubusercontent.com/BennyThadikaran/eod2_data

    CORRECTED: the JSON file has two top-level keys, not one flat map -
    "sym2isin" (symbol -> isin, what we want) and "isin2hist" (isin -> list
    of symbol-change records, e.g. renames/corporate actions - not needed
    here). Earlier version of this function read the whole file as if it
    were already the flat map, which put "sym2isin"/"isin2hist" themselves
    into the stock picker as fake symbols.
    """
    url = "https://raw.githubusercontent.com/BennyThadikaran/eod2_data/main/isin_symbol_map.json"
    import requests
    data = requests.get(url, timeout=15).json()
    return data["sym2isin"]


# ---------------------------------------------------------------------------
# 2. EOD2 PRICE / VOLUME / DELIVERY DATA
# ---------------------------------------------------------------------------

def load_eod2_data(symbol: str) -> pd.DataFrame:
    """
    Fetch daily OHLCV + delivery data for one NSE symbol.

    Confirmed columns (verified live on RELIANCE):
    Date, Open, High, Low, Close, Volume, Series, TOTAL_TRADES, QTY_PER_TRADE, DLV_QTY

    IMPORTANT (verified, not assumed): the filename must be lowercase,
    e.g. daily/reliance.csv - uppercase 404s.
    """
    url = f"https://raw.githubusercontent.com/BennyThadikaran/eod2_data/main/daily/{symbol.lower()}.csv"
    df = pd.read_csv(url, parse_dates=["Date"])
    return df


def calc_delivery_pct(df: pd.DataFrame) -> pd.Series:
    """Delivery % = DLV_QTY / Volume * 100, per your instruction (not raw DLV_QTY)."""
    return (df["DLV_QTY"] / df["Volume"].replace(0, np.nan)) * 100


def calc_relative_delivery(df: pd.DataFrame, lookback: int = 252, method: str = "zscore") -> pd.Series:
    """
    "Relative" delivery %, as requested - normalized against the stock's OWN history.
    This is the concrete formula behind VADM_t's H2 (delivery flow z-score) already
    named in your project memory - not a new concept, just its implementation.

    method="zscore"     -> (today's delivery% - rolling mean) / rolling std
    method="percentile" -> rolling percentile rank of today's delivery%, 0-100
    """
    dlv_pct = calc_delivery_pct(df)
    if method == "zscore":
        roll_mean = dlv_pct.rolling(lookback).mean()
        roll_std = dlv_pct.rolling(lookback).std()
        return (dlv_pct - roll_mean) / roll_std.replace(0, np.nan)
    elif method == "percentile":
        return dlv_pct.rolling(lookback).rank(pct=True) * 100
    else:
        raise ValueError("method must be 'zscore' or 'percentile'")


def calc_volume_percentile(df: pd.DataFrame, lookback: int = 60) -> pd.Series:
    """
    Rolling percentile rank of Volume over `lookback` trading days.
    lookback is a parameter, not hardcoded - wire this to a Streamlit slider,
    per our earlier agreement that lookback should be a UI toggle.
    """
    return df["Volume"].rolling(lookback).rank(pct=True) * 100


def calc_52wk_high_low(df: pd.DataFrame) -> tuple:
    """Returns (52wk_high, 52wk_low) using the most recent ~252 trading days."""
    recent = df.tail(252)
    return recent["High"].max(), recent["Low"].min()


def fetch_market_depth(symbol: str, download_folder: str = "./nse_data") -> dict:
    """
    Live order book (market depth) for exit planning.

    VERIFIED STRUCTURE (I fetched the package's own sample response and
    confirmed this): nse.quote(symbol)["orderBook"] contains buyPrice1-5 /
    buyQuantity1-5 (bid ladder) and sellPrice1-5 / sellQuantity1-5 (ask
    ladder), plus totalBuyQuantity / totalSellQuantity / lastPrice.

    STILL UNTESTED LIVE: I confirmed the *shape* from the package's sample
    file, not by actually calling nse.quote() against the live site (this
    sandbox can't reach nseindia.com). Run it yourself before trusting it.

    server=True for the same reason as fetch_promoter_holding - this runs
    on Streamlit Community Cloud, a server environment.
    """
    from nse import NSE

    with NSE(download_folder, server=True) as nse_client:
        data = nse_client.quote(symbol=symbol)

    ob = data.get("orderBook", {})
    bids = [
        {"price": ob.get(f"buyPrice{i}"), "qty": ob.get(f"buyQuantity{i}")}
        for i in range(1, 6)
    ]
    asks = [
        {"price": ob.get(f"sellPrice{i}"), "qty": ob.get(f"sellQuantity{i}")}
        for i in range(1, 6)
    ]
    return {
        "bids": bids,
        "asks": asks,
        "total_buy_qty": ob.get("totalBuyQuantity"),
        "total_sell_qty": ob.get("totalSellQuantity"),
        "last_price": ob.get("lastPrice"),
    }


def estimate_exit_price(bids: list, exit_quantity: int) -> dict:
    """
    Exiting a LONG position means SELLING, which hits the BID side of the
    book (the buyers), not the ask side - that's the part of "best exit
    rate" that's easy to get backwards, so this is deliberate, not a typo.

    Walks the bid ladder from best price down, filling `exit_quantity`
    cumulatively, to estimate the volume-weighted average fill price.

    Only 5 levels are available from the free NSE quote endpoint - if
    exit_quantity exceeds the visible depth, `unfilled_qty` will be > 0,
    meaning the real fill price would be worse than what's shown here
    (there's demand beyond what NSE's free quote exposes).
    """
    remaining = exit_quantity
    total_value = 0.0
    filled = 0

    for level in bids:
        if level["price"] is None or level["qty"] is None:
            continue
        take = min(remaining, level["qty"])
        total_value += take * level["price"]
        filled += take
        remaining -= take
        if remaining <= 0:
            break

    vwap = (total_value / filled) if filled else None
    return {
        "estimated_vwap_price": vwap,
        "filled_qty": filled,
        "unfilled_qty": max(remaining, 0),
        "depth_sufficient": remaining <= 0,
    }


# ---------------------------------------------------------------------------
# 3. TECHNICAL INDICATORS (for signal logic - separate from chart display,
#    which uses the lightweight-charts-v5 package's own Indicator classes)
# ---------------------------------------------------------------------------

def calc_sma(df: pd.DataFrame, window: int = 20) -> pd.Series:
    return df["Close"].rolling(window).mean()


def calc_rsi(df: pd.DataFrame, window: int = 14) -> pd.Series:
    delta = df["Close"].diff()
    gain = delta.clip(lower=0).rolling(window).mean()
    loss = (-delta.clip(upper=0)).rolling(window).mean()
    rs = gain / loss.replace(0, np.nan)
    return 100 - (100 / (1 + rs))


def calc_macd(df: pd.DataFrame, fast=12, slow=26, signal=9):
    ema_fast = df["Close"].ewm(span=fast, adjust=False).mean()
    ema_slow = df["Close"].ewm(span=slow, adjust=False).mean()
    macd_line = ema_fast - ema_slow
    signal_line = macd_line.ewm(span=signal, adjust=False).mean()
    return macd_line, signal_line, macd_line - signal_line


# ---------------------------------------------------------------------------
# 4. SCREENER EXCEL PARSING
#    Tested against your actual IRB_Infra_Devl-3.xlsx structure.
#    Real values live in "Data Sheet" - the other tabs (Profit & Loss,
#    Balance Sheet, Cash Flow) came back empty when read programmatically.
# ---------------------------------------------------------------------------

def parse_screener_excel(filepath: str) -> dict:
    """
    Parses a Screener.in single-company export (confirmed structure from
    your IRB Infra file). Finds rows by label text in column A, so it should
    survive minor row-position differences between different companies'
    exports - but this has ONLY been tested on the one file you gave me.
    Test it on a second company's export before trusting it broadly.
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb["Data Sheet"]

    rows = {}
    for row in ws.iter_rows():
        label = row[0].value
        if isinstance(label, str) and label.strip():
            rows[label.strip().rstrip(":")] = [c.value for c in row[1:]]

    def get(label):
        for key in rows:
            if key.upper().startswith(label.upper()):
                return rows[key]
        return None

    report_dates = get("Report Date")
    sales = get("Sales")
    net_profit = get("Net Profit")
    price = get("PRICE")
    adj_shares = get("Adjusted Equity Shares in Cr")
    borrowings = get("Borrowings")
    cash_bank = get("Cash & Bank")

    return {
        "years": report_dates,
        "sales": sales,
        "net_profit": net_profit,
        "price": price,
        "adjusted_shares_cr": adj_shares,
        "borrowings": borrowings,
        "cash_and_bank": cash_bank,
    }


def calc_pe_per_year(price: list, net_profit: list, adj_shares_cr: list) -> list:
    """
    PE_year = Price_year / EPS_year, where EPS_year = Net Profit / Adjusted Shares.
    Net Profit is in Cr, Adjusted Shares is in Cr, so units cancel to give EPS in Rs.

    NOTE: I'm inferring "PRICE" row = FY-end close. Sanity-check one year against
    Screener.in's own displayed PE chart for the same company before trusting this.
    """
    pe_values = []
    for p, np_, shares in zip(price, net_profit, adj_shares_cr):
        if None in (p, np_, shares) or shares == 0 or np_ == 0:
            pe_values.append(None)
            continue
        eps = np_ / shares
        pe_values.append(p / eps if eps else None)
    return pe_values


def calc_market_cap_per_year(price: list, adj_shares_cr: list) -> list:
    """Market Cap_year = Price_year * Adjusted Shares_year (Cr). Same building
    block calc_ev_per_year uses internally - pulled out so the UI can show it
    as its own card without recomputing the formula separately."""
    mcap = []
    for p, shares in zip(price, adj_shares_cr):
        mcap.append(p * shares if (p is not None and shares is not None) else None)
    return mcap


def calc_ev_per_year(price: list, adj_shares_cr: list, borrowings: list, cash_bank: list) -> list:
    """
    EV_year = MarketCap_year + Borrowings_year - Cash&Bank_year
    MarketCap_year = Price_year * Adjusted Shares_year (in Cr, so result is in Cr).
    This is the simplified EV formula - no minority interest / preferred equity
    adjustment, which more complete EV formulas sometimes include.
    """
    ev_values = []
    for p, shares, debt, cash in zip(price, adj_shares_cr, borrowings, cash_bank):
        if None in (p, shares, debt, cash):
            ev_values.append(None)
            continue
        market_cap = p * shares
        ev_values.append(market_cap + debt - cash)
    return ev_values


def calc_revenue_growth(sales: list) -> list:
    """YoY revenue growth %, first year is None (no prior year to compare)."""
    growth = [None]
    for i in range(1, len(sales)):
        if sales[i - 1] in (None, 0) or sales[i] is None:
            growth.append(None)
        else:
            growth.append((sales[i] - sales[i - 1]) / sales[i - 1] * 100)
    return growth


def estimate_capex(net_block: list, depreciation: list) -> list:
    """
    APPROXIMATION, not a real reported figure - flagged per your confirmation
    that a labeled approximation is acceptable.
    Capex ~= change in Net Block + Depreciation for that year.
    This is a common shorthand but will misstate capex in years with asset
    write-offs, revaluations, or disposals - it doesn't account for those.
    """
    capex = [None]
    for i in range(1, len(net_block)):
        if net_block[i] is None or net_block[i - 1] is None or depreciation[i] is None:
            capex.append(None)
        else:
            capex.append((net_block[i] - net_block[i - 1]) + depreciation[i])
    return capex


# ---------------------------------------------------------------------------
# 5. PROMOTER HOLDING (via `nse` package)
#    UNTESTED - this sandbox cannot reach nseindia.com. Written to match the
#    documented nse.shareholding() signature exactly. Run it yourself first.
# ---------------------------------------------------------------------------

def fetch_promoter_holding(symbol: str, download_folder: str = "./nse_data") -> dict:
    """
    Returns the most recent quarter's promoter holding %.
    Per the docs you pasted: nse.shareholding() returns a list of dicts,
    most recent quarter first, with 'pr_and_prgrp' as the promoter+group field.

    Pledge is NOT a documented field here - confirm by printing the full
    dict for one symbol before assuming it's absent or present.
    """
    from nse import NSE  # import here so the rest of this module works even
                          # if `nse` isn't installed in some environments

    # server=True because this now runs on Streamlit Community Cloud, not your
    # own machine - the nse package's own docs say server=True is needed for
    # cloud/server environments (and it switches the HTTP backend to httpx).
    with NSE(download_folder, server=True) as nse_client:
        data = nse_client.shareholding(symbol=symbol, index="equities")

    if not data:
        return {"symbol": symbol, "promoter_holding_pct": None, "as_of": None}

    latest = data[0]
    return {
        "symbol": symbol,
        "promoter_holding_pct": latest.get("pr_and_prgrp"),
        "as_of": latest.get("date"),
        "raw": latest,  # keep the full record so you can inspect other fields
    }


# ---------------------------------------------------------------------------
# 6. ALPHA WHITE SIGNAL LOGIC
#    The comparison itself is complete. Where sector_avg_pe comes from is
#    still open - see note at bottom of file.
# ---------------------------------------------------------------------------

def alpha_white_signal(stock_pe: float, sector_avg_pe: float,
                        current_volume_percentile: float,
                        volume_percentile_threshold: float = 80.0) -> str:
    """
    Buy: PE < sector average PE  AND  Volume percentile > threshold
    Sell: otherwise
    (Your confirmed definition from earlier in this project.)

    sector_avg_pe has no resolved data source yet (see below) - caller must
    supply it; this function does not fetch or guess it.
    """
    if stock_pe is None or sector_avg_pe is None or current_volume_percentile is None:
        return "INSUFFICIENT_DATA"

    pe_low = stock_pe < sector_avg_pe
    volume_high = current_volume_percentile > volume_percentile_threshold

    return "BUY" if (pe_low and volume_high) else "SELL"


# ---------------------------------------------------------------------------
# STILL OPEN - NOT BUILT, NOT GUESSED:
#
# 1. sector_avg_pe data source. After the Excel turned out to be a
#    single-company export with no Industry/Sector column, this reverted to
#    an unresolved question. Options still on the table from earlier in this
#    conversation: (a) a manual stock->sector mapping table you maintain,
#    (b) NSE sectoral index PE via a library, if you want to revisit that.
#    Pick one before Alpha White can run end-to-end.
#
# 2. Alpha Black / VADM_t. No formula f(), no H4. Your own project memory
#    flags this explicitly as a CEO-level decision - nothing here invents it.
# ---------------------------------------------------------------------------
