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
  - Alpha White PE condition......Switched to SELF-RELATIVE (vs own historical
                                  PE range, percentile-based, toggleable 20/80
                                  defaults) instead of sector-average - that
                                  data source was never resolved, so this
                                  sidesteps it rather than waiting. Coarse
                                  resolution caveat: only ~10 annual PE points
                                  to percentile against (see function docstring).
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

    NOTE: this universe includes symbols NSE has EVER had, including
    delisted/merged ones (e.g. HDFC, merged into HDFCBANK in 2023) - picking
    one of those will 404 in load_eod2_data(). Use check_symbol_status()
    below to explain why, using the isin2hist data this same file provides.
    """
    url = "https://raw.githubusercontent.com/BennyThadikaran/eod2_data/main/isin_symbol_map.json"
    import requests
    data = requests.get(url, timeout=15).json()
    return data["sym2isin"]


def check_symbol_status(symbol: str) -> dict:
    """
    Looks up a symbol's own history record (from_date/to_date/action) using
    the same isin_symbol_map.json data fetch_stock_universe() reads, so we
    can explain a 404 instead of just showing a raw HTTP error - e.g.
    "HDFC" merged into HDFCBANK on 2023-07-12, confirmed via this exact
    lookup (from_date 2011-06-22, to_date 2023-07-12, action: None -
    NSE's own record doesn't always fill in 'action' even for known
    mergers, so we can confirm the stop-date but not always the reason).
    """
    import requests
    url = "https://raw.githubusercontent.com/BennyThadikaran/eod2_data/main/isin_symbol_map.json"
    data = requests.get(url, timeout=15).json()

    isin = data["sym2isin"].get(symbol)
    if isin is None:
        return {"found": False, "message": f"'{symbol}' isn't in the known NSE symbol list at all."}

    history = data["isin2hist"].get(isin, [])
    if not history:
        return {"found": True, "history": [], "message": "No history record found - unclear why the data fetch failed."}

    latest = history[0]
    return {
        "found": True,
        "history": history,
        "from_date": latest.get("from_date"),
        "to_date": latest.get("to_date"),
        "action": latest.get("action"),
        "message": (
            f"'{symbol}' was active from {latest.get('from_date')} to {latest.get('to_date')}. "
            f"If that end date is in the past, it likely stopped trading under this symbol "
            f"(delisting, merger, or rename) - search for a successor symbol if applicable."
        ),
    }


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


def calc_volume_direction(df: pd.DataFrame) -> pd.Series:
    """
    Classifies which way each day's volume leaned using Close Location
    Value (CLV) - where the Close sat within that day's own High-Low range.
    Per your instruction (the "more precise" option over simple Close-vs-
    prev-Close):

        CLV = ((Close - Low) - (High - Close)) / (High - Low)

    +1 = closed at the day's high (pure buying pressure)
    -1 = closed at the day's low (pure selling pressure)
     0 = closed at the exact midpoint

    This is the standard building block behind Accumulation/Distribution-
    style indicators - not something invented for this project, just the
    well-known formula, applied here for your specific purpose.
    Zero-range days (circuit-locked, High==Low) are guarded to avoid
    divide-by-zero - they come back as NaN, correctly excluded rather than
    silently treated as buy or sell.
    """
    day_range = (df["High"] - df["Low"]).replace(0, np.nan)
    clv = ((df["Close"] - df["Low"]) - (df["High"] - df["Close"])) / day_range
    return clv


def calc_volume_regime(df: pd.DataFrame, lookback: int = 60,
                        percentile_threshold: float = 80.0) -> dict:
    """
    Combines volume percentile (how unusual today's volume is vs its own
    history) with CLV (which direction that volume leaned) into a single
    classification for TODAY:

      HEAVY_BUY  - volume above the percentile threshold AND CLV > 0
      HEAVY_SELL - volume above the percentile threshold AND CLV < 0
      NORMAL     - volume not unusual, or direction unclear (CLV == 0)

    Returns the raw percentile/CLV values too, so the UI can show them
    rather than just the label.
    """
    vol_pctile_series = calc_volume_percentile(df, lookback=lookback)
    clv_series = calc_volume_direction(df)

    current_vol_pctile = vol_pctile_series.iloc[-1]
    current_clv = clv_series.iloc[-1]

    if pd.isna(current_vol_pctile) or pd.isna(current_clv):
        return {"regime": "INSUFFICIENT_DATA", "volume_percentile": None, "clv": None}

    is_heavy = current_vol_pctile > percentile_threshold

    if is_heavy and current_clv > 0:
        regime = "HEAVY_BUY"
    elif is_heavy and current_clv < 0:
        regime = "HEAVY_SELL"
    else:
        regime = "NORMAL"

    return {"regime": regime, "volume_percentile": float(current_vol_pctile), "clv": float(current_clv)}


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


def parse_multiple_screener_excels(symbol_filepath_pairs: list) -> dict:
    """
    Batch version of parse_screener_excel - reduces the tedium of your
    chosen approach (manual Screener uploads, one per stock) to "upload N
    files at once" rather than repeating the single-stock flow N times.
    Reuses parse_screener_excel exactly, already verified against your
    real IRB file - this is just a loop around it, nothing new to trust.

    symbol_filepath_pairs: list of (symbol, filepath) tuples.
    Returns {"parsed": {symbol: parsed_dict, ...}, "failures": [(symbol, error), ...]}
    - failures don't stop the batch, so one bad file doesn't block the rest.
    """
    parsed = {}
    failures = []
    for symbol, filepath in symbol_filepath_pairs:
        try:
            parsed[symbol] = parse_screener_excel(filepath)
        except Exception as e:
            failures.append((symbol, str(e)))
    return {"parsed": parsed, "failures": failures}


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
#    PE side now self-relative (vs own history), not sector-relative -
#    sidesteps the sector-data-source problem instead of waiting on it.
# ---------------------------------------------------------------------------

def calc_current_pe(cmp: float, latest_net_profit_cr: float, latest_adj_shares_cr: float):
    """
    "Current PE" - today's LIVE market price divided by the most recent
    annual EPS available. Not a true TTM PE (no quarterly EPS wired in yet),
    so this can lag if results have moved a lot since the last annual report -
    a common simplification, not a precise trailing-twelve-month figure.
    """
    if latest_net_profit_cr in (None, 0) or latest_adj_shares_cr in (None, 0):
        return None
    eps = latest_net_profit_cr / latest_adj_shares_cr
    return cmp / eps if eps else None


def calc_relative_pe_regime(current_pe, historical_pe_series: list,
                             low_percentile: float = 20, high_percentile: float = 80) -> dict:
    """
    "Relative PE" per your instruction - NOT vs sector (that data source was
    never resolved), but vs the stock's OWN historical PE range. Same
    self-relative logic as the volume percentile toggle, thresholds
    adjustable, default 20/80 as you specified.

    CAVEAT, flagged plainly: your screener export gives ~10 annual PE
    points only. A percentile off 10 points is coarse - each percentile
    step covers roughly one data point, not a smooth distribution. This
    works but is low-resolution. A sharper version would need a daily PE
    series (rolling TTM EPS from quarterly results, e.g. via
    nse.results_comparison() - untested from this sandbox) - bigger scope,
    not built here unless you want it.
    """
    clean_hist = [p for p in historical_pe_series if p is not None]
    if not clean_hist or current_pe is None:
        return {"regime": "INSUFFICIENT_DATA", "low_threshold": None, "high_threshold": None}

    low_thresh = float(np.percentile(clean_hist, low_percentile))
    high_thresh = float(np.percentile(clean_hist, high_percentile))

    if current_pe < low_thresh:
        regime = "LOW"
    elif current_pe > high_thresh:
        regime = "HIGH"
    else:
        regime = "NEUTRAL"

    return {"regime": regime, "low_threshold": low_thresh, "high_threshold": high_thresh}


def alpha_white_signal(pe_regime: str, volume_regime: str) -> str:
    """
    UPDATED per your clarification - this is the real definition now:

      BUY:  PE regime LOW   AND  Volume regime HEAVY_BUY
      SELL: PE regime HIGH  AND  Volume regime HEAVY_SELL
      HOLD: everything else - PE NEUTRAL, or PE/Volume direction mismatched
            (e.g. PE cheap but volume is HEAVY_SELL, or PE expensive but
            volume is HEAVY_BUY). You explicitly confirmed this should NOT
            be binary BUY/SELL anymore - HOLD is a real third state.

    Kept separate from INSUFFICIENT_DATA, which means we're missing the
    data to judge at all (different from "judged, and it's a hold").
    """
    if pe_regime in (None, "INSUFFICIENT_DATA") or volume_regime in (None, "INSUFFICIENT_DATA"):
        return "INSUFFICIENT_DATA"

    if pe_regime == "LOW" and volume_regime == "HEAVY_BUY":
        return "BUY"
    if pe_regime == "HIGH" and volume_regime == "HEAVY_SELL":
        return "SELL"
    return "HOLD"


# ---------------------------------------------------------------------------
# 7. BACKTEST - runs the Alpha White signal across full price history,
#    with zero lookahead bias in the PE percentile thresholds.
# ---------------------------------------------------------------------------

def run_alpha_white_backtest(eod_df: pd.DataFrame, fy_dates: list, fy_price: list,
                              fy_net_profit: list, fy_adj_shares: list,
                              volume_lookback: int = 60, volume_pctile_threshold: float = 80.0,
                              pe_low_pctile: float = 20, pe_high_pctile: float = 80,
                              reporting_lag_days: int = 60, min_history_points: int = 3) -> pd.DataFrame:
    """
    Runs Alpha White's BUY/SELL/HOLD signal across every day in eod_df,
    with NO LOOKAHEAD BIAS - per your explicit requirement.

    How the no-lookahead guarantee works:
      - Each FY's EPS is only treated as "known" starting reporting_lag_days
        AFTER its FY-end date (default 60 days) - NOT the FY-end date itself.
        FLAGGED ASSUMPTION: Screener's Data Sheet only gives FY-END dates,
        not actual results-announcement dates. Most Indian companies report
        annual results roughly 1-2 months after FY end, so 60 days is a
        conservative buffer against using EPS before it was plausibly
        public. If you know the real announcement dates, this should use
        those instead - ask if you want that precision.
      - On any given day, the PE percentile thresholds (20th/80th) are
        computed ONLY from FY-end PE values that were already "known" as of
        that day - never using a threshold informed by future PE data.
      - Volume regime (percentile + CLV) was ALREADY point-in-time safe
        (rolling calculations) - no change needed there.
      - Days before `min_history_points` annual PE values have become
        available return INSUFFICIENT_DATA rather than a meaningless
        percentile off 1-2 data points.

    Returns a DataFrame: Date, Close, Volume, current_pe, pe_low_threshold,
    pe_high_threshold, pe_regime, volume_percentile, clv, volume_regime, signal
    - one row per trading day, so you can see exactly how the stock was
      moving and what volume looked like at every point, not just on
      signal days.
    """
    # Per-FY EPS + own annual PE (the "population" for future percentiles)
    annual_pe = calc_pe_per_year(fy_price, fy_net_profit, fy_adj_shares)
    fy_eps = [
        (npft / sh) if (npft not in (None, 0) and sh not in (None, 0)) else None
        for npft, sh in zip(fy_net_profit, fy_adj_shares)
    ]

    fy_frame = pd.DataFrame({
        "fy_end": pd.to_datetime(fy_dates),
        "eps": fy_eps,
        "annual_pe": annual_pe,
    }).dropna(subset=["eps"]).sort_values("fy_end").reset_index(drop=True)
    fy_frame["available_date"] = fy_frame["fy_end"] + pd.Timedelta(days=reporting_lag_days)

    daily = eod_df[["Date", "Close", "High", "Low", "Volume"]].copy().sort_values("Date").reset_index(drop=True)

    # merge_asof: for each trading day, attach the most recently AVAILABLE
    # EPS as of that day (backward-looking only - this is the crux of the
    # no-lookahead guarantee for the PE numerator).
    eps_lookup = fy_frame[["available_date", "eps"]].rename(columns={"available_date": "Date"})
    merged = pd.merge_asof(daily, eps_lookup, on="Date", direction="backward")
    merged["current_pe"] = merged["Close"] / merged["eps"]

    # Percentile thresholds: expanding population of annual PE values,
    # only using ones already "available" as of each date - looped over
    # FY boundaries (only ~10 iterations), NOT over every daily row, so
    # this stays fast.
    merged["pe_low_threshold"] = np.nan
    merged["pe_high_threshold"] = np.nan

    fy_avail = fy_frame.dropna(subset=["annual_pe"]).sort_values("available_date").reset_index(drop=True)
    for i in range(len(fy_avail)):
        cutoff = fy_avail.loc[i, "available_date"]
        pe_population = fy_avail.loc[:i, "annual_pe"].dropna().tolist()
        mask = merged["Date"] >= cutoff
        if i + 1 < len(fy_avail):
            mask &= merged["Date"] < fy_avail.loc[i + 1, "available_date"]
        if len(pe_population) >= min_history_points:
            merged.loc[mask, "pe_low_threshold"] = float(np.percentile(pe_population, pe_low_pctile))
            merged.loc[mask, "pe_high_threshold"] = float(np.percentile(pe_population, pe_high_pctile))

    def _pe_regime(row):
        if pd.isna(row["pe_low_threshold"]) or pd.isna(row["current_pe"]):
            return "INSUFFICIENT_DATA"
        if row["current_pe"] < row["pe_low_threshold"]:
            return "LOW"
        if row["current_pe"] > row["pe_high_threshold"]:
            return "HIGH"
        return "NEUTRAL"

    merged["pe_regime"] = merged.apply(_pe_regime, axis=1)

    # Volume regime - reuse the already-rolling, already-safe functions.
    eod_sorted = eod_df.sort_values("Date").reset_index(drop=True)
    merged["volume_percentile"] = calc_volume_percentile(eod_sorted, lookback=volume_lookback).values
    merged["clv"] = calc_volume_direction(eod_sorted).values

    def _volume_regime(row):
        if pd.isna(row["volume_percentile"]) or pd.isna(row["clv"]):
            return "INSUFFICIENT_DATA"
        is_heavy = row["volume_percentile"] > volume_pctile_threshold
        if is_heavy and row["clv"] > 0:
            return "HEAVY_BUY"
        if is_heavy and row["clv"] < 0:
            return "HEAVY_SELL"
        return "NORMAL"

    merged["volume_regime"] = merged.apply(_volume_regime, axis=1)
    merged["signal"] = [alpha_white_signal(pe, vol) for pe, vol in
                         zip(merged["pe_regime"], merged["volume_regime"])]

    return merged[["Date", "Close", "Volume", "current_pe", "pe_low_threshold", "pe_high_threshold",
                    "pe_regime", "volume_percentile", "clv", "volume_regime", "signal"]]


def summarize_backtest_signals(backtest_df: pd.DataFrame, holding_days: int = 20) -> pd.DataFrame:
    """
    For every BUY/SELL day the backtest found, computes what ACTUALLY
    happened to price over the next `holding_days` trading days - per your
    instruction to see real results, not just signal counts.
    Rows near the very end of the data (not enough future days left) get
    forward_return_pct = None rather than a wrong number.
    """
    df = backtest_df.reset_index(drop=True)
    signal_rows = df[df["signal"].isin(["BUY", "SELL"])].copy()

    forward_returns = []
    for idx in signal_rows.index:
        entry_price = df.loc[idx, "Close"]
        exit_idx = idx + holding_days
        if exit_idx < len(df):
            exit_price = df.loc[exit_idx, "Close"]
            forward_returns.append((exit_price - entry_price) / entry_price * 100)
        else:
            forward_returns.append(None)

    signal_rows["forward_return_pct"] = forward_returns
    return signal_rows[["Date", "Close", "signal", "pe_regime", "volume_regime", "forward_return_pct"]]


def grid_search_alpha_white_thresholds(
    stock_configs: list,
    pe_low_pctile_grid: list = None,
    pe_high_pctile_grid: list = None,
    volume_pctile_threshold_grid: list = None,
    volume_lookback: int = 60,
    holding_days: int = 20,
    min_signals_required: int = 10,
) -> pd.DataFrame:
    """
    Grid search over Alpha White's threshold parameters, POOLED across
    multiple stocks, to find genuinely best-performing thresholds instead
    of the arbitrary 20/80/80 defaults we started with - per your decision
    to tune before adding risk parameters.

    stock_configs: list of dicts, one per stock, each with keys:
      symbol, eod_df, fy_dates, fy_price, fy_net_profit, fy_adj_shares
    (fy_* fields come from parse_screener_excel() for that stock - PE
    thresholds need Screener fundamental data, unlike volume thresholds
    which only need EOD2 and could scale to any number of stocks freely.)

    For each parameter combination, runs run_alpha_white_backtest per stock,
    POOLS all BUY signals together across stocks (and SELL separately),
    then reports average forward return, win rate, and signal count.

    Combinations producing fewer than min_signals_required TOTAL pooled
    signals are flagged unreliable=False - not enough observations to
    trust that combination's average, same discipline as everywhere else
    in this file. A combination with a great-looking average return off
    3 signals is noise, not a finding.

    Defaults for the grids (if None) are deliberately centered on the
    20/80/80 starting point, spanning a reasonable neighborhood - not
    an exhaustive search of all possible values.
    """
    if pe_low_pctile_grid is None:
        pe_low_pctile_grid = [10, 15, 20, 25, 30]
    if pe_high_pctile_grid is None:
        pe_high_pctile_grid = [70, 75, 80, 85, 90]
    if volume_pctile_threshold_grid is None:
        volume_pctile_threshold_grid = [70, 75, 80, 85, 90, 95]

    results = []

    for pe_low in pe_low_pctile_grid:
        for pe_high in pe_high_pctile_grid:
            if pe_low >= pe_high:
                continue  # nonsensical combination - low threshold above high threshold

            for vol_thresh in volume_pctile_threshold_grid:
                all_buy_returns, all_sell_returns = [], []

                for config in stock_configs:
                    try:
                        bt = run_alpha_white_backtest(
                            config["eod_df"], config["fy_dates"], config["fy_price"],
                            config["fy_net_profit"], config["fy_adj_shares"],
                            volume_lookback=volume_lookback, volume_pctile_threshold=vol_thresh,
                            pe_low_pctile=pe_low, pe_high_pctile=pe_high,
                        )
                        summary = summarize_backtest_signals(bt, holding_days=holding_days)
                        all_buy_returns.extend(
                            summary[summary["signal"] == "BUY"]["forward_return_pct"].dropna().tolist()
                        )
                        all_sell_returns.extend(
                            summary[summary["signal"] == "SELL"]["forward_return_pct"].dropna().tolist()
                        )
                    except Exception:
                        continue  # one stock's failure shouldn't kill the whole grid point

                buy_avg = float(np.mean(all_buy_returns)) if all_buy_returns else None
                buy_win_rate = (
                    float(sum(1 for r in all_buy_returns if r > 0) / len(all_buy_returns) * 100)
                    if all_buy_returns else None
                )
                sell_avg = float(np.mean(all_sell_returns)) if all_sell_returns else None
                sell_win_rate = (
                    float(sum(1 for r in all_sell_returns if r < 0) / len(all_sell_returns) * 100)
                    if all_sell_returns else None
                )

                results.append({
                    "pe_low_pctile": pe_low, "pe_high_pctile": pe_high,
                    "volume_pctile_threshold": vol_thresh,
                    "n_buy_signals": len(all_buy_returns), "buy_avg_return_pct": buy_avg,
                    "buy_win_rate_pct": buy_win_rate,
                    "n_sell_signals": len(all_sell_returns), "sell_avg_return_pct": sell_avg,
                    "sell_win_rate_pct": sell_win_rate,
                    "reliable": (len(all_buy_returns) >= min_signals_required
                                 and len(all_sell_returns) >= min_signals_required),
                })

    return pd.DataFrame(results)


def build_signal_markers(signal_rows_df: pd.DataFrame) -> list:
    """
    Converts backtest signal rows into the marker format the chart package
    expects (confirmed from its own source): time/position/color/shape/text.
    Feed the output straight into PriceIndicator(markers=...).
    """
    markers = []
    for _, row in signal_rows_df.iterrows():
        if row["signal"] == "BUY":
            markers.append({
                "time": row["Date"].strftime("%Y-%m-%d") if hasattr(row["Date"], "strftime") else str(row["Date"]),
                "position": "belowBar", "color": "#26a69a", "shape": "arrowUp", "text": "BUY",
            })
        elif row["signal"] == "SELL":
            markers.append({
                "time": row["Date"].strftime("%Y-%m-%d") if hasattr(row["Date"], "strftime") else str(row["Date"]),
                "position": "aboveBar", "color": "#ef5350", "shape": "arrowDown", "text": "SELL",
            })
    return markers


# ---------------------------------------------------------------------------
# 8. DELIVERY COMPOSITE MODEL - "Altman-style" per your request: weights are
#    FIT from real forward-return data via regression, not hand-assigned.
#    This is a step toward VADM_t's Accumulation/Distribution axis, not
#    VADM_t itself - the combination with the PE axis still needs the
#    formula/H4 that's a CEO-level decision (see note at bottom of file).
# ---------------------------------------------------------------------------

def build_delivery_features(df: pd.DataFrame, lookback: int = 252) -> pd.DataFrame:
    """
    Three GENUINELY DISTINCT delivery-related angles, analogous to Altman's
    5 ratios being different dimensions of financial health - not just
    transforms of the same number (delivery% and its z-score would be too
    similar to count as separate "ratios"):

      delivery_zscore  - how unusual TODAY's delivery% is vs its own
                          rolling history (already had this)
      delivery_trend   - short-term momentum: 5-day average delivery% minus
                          20-day average - is delivery% itself heating up
                          or cooling down recently
      delivery_price_character - rolling average of (CLV * delivery%) -
                          captures whether HIGH-delivery days have tended
                          to close strong (near the day's high) or weak
                          (near the day's low). Reuses the CLV formula
                          already built for Alpha White's volume direction.
    """
    dlv_pct = calc_delivery_pct(df)
    clv = calc_volume_direction(df)

    delivery_zscore = calc_relative_delivery(df, lookback=lookback, method="zscore")
    delivery_trend = dlv_pct.rolling(5).mean() - dlv_pct.rolling(20).mean()
    delivery_price_character = (clv * dlv_pct).rolling(20).mean()

    return pd.DataFrame({
        "delivery_zscore": delivery_zscore,
        "delivery_trend": delivery_trend,
        "delivery_price_character": delivery_price_character,
    })


def fit_delivery_score_model(df: pd.DataFrame, forward_days: int = 20,
                              lookback: int = 252, test_fraction: float = 0.3) -> dict:
    """
    Fits weights for the 3 delivery features against ACTUAL forward returns,
    via OLS regression - the genuinely-Altman-style part: weights come from
    data, not from feel.

    HONESTY SAFEGUARDS built in, not optional:
      - Chronological train/test split (NOT random shuffling) - a random
        split would let the model "see" data from both before and after any
        given test point, which is its own form of lookahead leakage.
      - Reports OUT-OF-SAMPLE R² as the real number to trust, not just
        in-sample fit (in-sample R² on a handful of features will almost
        always look better than it deserves to).
      - Reports p-values per weight - a weight with a high p-value means
        "this coefficient is statistically indistinguishable from noise",
        and should not be trusted as a real signal even though the
        regression will still spit out SOME number for it.

    REAL LIMITATION, not hidden: this fits on ONE stock's own history.
    A model fit on a single stock may just be capturing that stock's
    idiosyncrasies, not a genuine generalizable pattern - a more robust
    version would pool data across many stocks. Not built here; ask if
    you want that scaled up.
    """
    import statsmodels.api as sm
    from sklearn.metrics import r2_score

    features = build_delivery_features(df, lookback=lookback)
    close = df["Close"].reset_index(drop=True)
    forward_return = (close.shift(-forward_days) - close) / close * 100

    data = features.reset_index(drop=True).copy()
    data["forward_return"] = forward_return.values
    data = data.dropna()

    if len(data) < 50:
        return {"error": f"Only {len(data)} clean overlapping rows after dropping NaNs - "
                          f"too few to fit reliably (need at least 50). Try a shorter "
                          f"lookback or a stock with longer price history."}

    split_idx = int(len(data) * (1 - test_fraction))
    train, test = data.iloc[:split_idx], data.iloc[split_idx:]
    feature_cols = ["delivery_zscore", "delivery_trend", "delivery_price_character"]

    X_train = sm.add_constant(train[feature_cols])
    y_train = train["forward_return"]
    model = sm.OLS(y_train, X_train).fit()

    X_test = sm.add_constant(test[feature_cols], has_constant="add")
    y_pred = model.predict(X_test)
    out_of_sample_r2 = r2_score(test["forward_return"], y_pred)

    return {
        "weights": model.params.to_dict(),
        "p_values": model.pvalues.to_dict(),
        "in_sample_r2": float(model.rsquared),
        "out_of_sample_r2": float(out_of_sample_r2),
        "n_train": len(train),
        "n_test": len(test),
    }


def fit_delivery_score_model_pooled(symbols: list, forward_days: int = 20, lookback: int = 252,
                                     test_fraction: float = 0.3, min_rows_per_stock: int = 300) -> dict:
    """
    Multi-stock version of fit_delivery_score_model - pools many stocks'
    data together instead of fitting on one, per your instruction, to
    address the overfitting the single-stock version showed on IRB
    (out-of-sample R² was -1.95 there).

    SPLIT METHODOLOGY - different from the single-stock version and
    important to get right: splits by an actual CALENDAR DATE cutoff
    applied across ALL pooled stocks together, not by row-position within
    each stock. A per-stock positional split would risk a stock with a
    shorter history leaking its "future" rows into training while another
    stock's genuinely earlier rows end up in test - the calendar-date cut
    avoids that.

    STILL A REAL LIMITATION, not hidden: pooled stocks aren't fully
    independent observations - broad market-wide moves hit many stocks on
    the same days. This reduces single-stock overfitting but doesn't make
    the rows truly i.i.d. the way genuinely separate companies would be.

    Any symbol that fails to fetch (delisted, network issue, etc.) is
    skipped and reported in fetch_failures rather than crashing the whole
    fit - useful given some symbols WILL be dead tickers (see check_symbol_status).
    """
    import statsmodels.api as sm
    from sklearn.metrics import r2_score

    all_rows = []
    fetch_failures = []

    for symbol in symbols:
        try:
            df = load_eod2_data(symbol)
            if len(df) < min_rows_per_stock:
                fetch_failures.append((symbol, f"only {len(df)} rows, below min_rows_per_stock"))
                continue
            features = build_delivery_features(df, lookback=lookback)
            close = df["Close"].reset_index(drop=True)
            forward_return = (close.shift(-forward_days) - close) / close * 100

            stock_data = features.reset_index(drop=True).copy()
            stock_data["forward_return"] = forward_return.values
            stock_data["Date"] = df["Date"].reset_index(drop=True).values
            stock_data["symbol"] = symbol
            all_rows.append(stock_data)
        except Exception as e:
            fetch_failures.append((symbol, str(e)))

    if not all_rows:
        return {"error": "Could not fetch usable data for any given symbol.", "fetch_failures": fetch_failures}

    pooled = pd.concat(all_rows, ignore_index=True).dropna(
        subset=["delivery_zscore", "delivery_trend", "delivery_price_character", "forward_return"]
    )

    if len(pooled) < 200:
        return {"error": f"Only {len(pooled)} clean pooled rows across all symbols - too few to "
                          f"fit reliably.", "fetch_failures": fetch_failures}

    pooled = pooled.sort_values("Date").reset_index(drop=True)
    split_idx = int(len(pooled) * (1 - test_fraction))
    split_date = pooled.loc[split_idx, "Date"]

    train = pooled[pooled["Date"] < split_date]
    test = pooled[pooled["Date"] >= split_date]
    feature_cols = ["delivery_zscore", "delivery_trend", "delivery_price_character"]

    X_train = sm.add_constant(train[feature_cols])
    model = sm.OLS(train["forward_return"], X_train).fit()

    X_test = sm.add_constant(test[feature_cols], has_constant="add")
    y_pred = model.predict(X_test)
    out_of_sample_r2 = r2_score(test["forward_return"], y_pred)

    return {
        "weights": model.params.to_dict(),
        "p_values": model.pvalues.to_dict(),
        "in_sample_r2": float(model.rsquared),
        "out_of_sample_r2": float(out_of_sample_r2),
        "n_train": len(train),
        "n_test": len(test),
        "n_stocks_used": int(pooled["symbol"].nunique()),
        "split_date": str(split_date),
        "fetch_failures": fetch_failures,
    }


def calc_delivery_composite_score(df: pd.DataFrame, weights: dict, lookback: int = 252) -> pd.Series:
    """
    Applies FITTED weights (from fit_delivery_score_model) to produce a
    single composite score per day - the actual "Altman-style" number.
    weights dict must have keys: const, delivery_zscore, delivery_trend,
    delivery_price_character (exactly what fit_delivery_score_model returns).
    """
    features = build_delivery_features(df, lookback=lookback)
    score = (
        weights.get("const", 0)
        + weights.get("delivery_zscore", 0) * features["delivery_zscore"]
        + weights.get("delivery_trend", 0) * features["delivery_trend"]
        + weights.get("delivery_price_character", 0) * features["delivery_price_character"]
    )
    return score


# ---------------------------------------------------------------------------
# 9. VADM_t - your exact spec: V (sector-relative valuation), D (delivery
#    Z-score composite), interaction term V*D, sector-grouped OLS.
#    STATUS: D is fully buildable now (no sector dependency). V needs the
#    sector mapping (built below, untested against live NSE) AND per-stock
#    EPS at scale - that second piece is still an open question, see the
#    note at the bottom of this section before assuming V/the regression
#    engine can run end-to-end.
# ---------------------------------------------------------------------------

def calc_delivery_score_D(df: pd.DataFrame, lookback: int = 20) -> pd.Series:
    """
    VADM_t's Delivery Score D, exactly per your spec:
      - 20-day rolling Z-score of ABSOLUTE Delivery Volume (DLV_QTY) -
        different from delivery %, this is raw size, not a ratio.
      - 20-day rolling Z-score of Deliverable % (DLV_QTY / Volume) - the
        "conviction" measure, reuses calc_relative_delivery's method.
      Combined into one composite D via simple unweighted average - a
      defensible default since the later regression's Beta_2 determines
      how much D as a WHOLE matters; I'm not privileging one sub-score
      over the other inside D itself without a reason to.

    Zero lookahead concern here: rolling z-scores are already point-in-time
    safe by construction (only look backward), same as the volume
    percentile functions already verified elsewhere in this file.
    """
    dlv_qty = df["DLV_QTY"]
    roll_mean_qty = dlv_qty.rolling(lookback).mean()
    roll_std_qty = dlv_qty.rolling(lookback).std()
    delivery_volume_zscore = (dlv_qty - roll_mean_qty) / roll_std_qty.replace(0, np.nan)

    deliverable_pct_zscore = calc_relative_delivery(df, lookback=lookback, method="zscore")

    D = (delivery_volume_zscore + deliverable_pct_zscore) / 2
    return D


def calc_pointintime_pe_series(eod_df: pd.DataFrame, fy_dates: list, fy_price: list,
                                fy_net_profit: list, fy_adj_shares: list,
                                reporting_lag_days: int = 60) -> pd.DataFrame:
    """
    Point-in-time-safe daily "current PE" for ONE stock - the same
    no-lookahead mechanism already proven inside run_alpha_white_backtest,
    pulled out here as its own reusable function since VADM_t's V-score
    needs this per-stock, not just for a single-stock live signal.
    Returns a DataFrame with columns: Date, current_pe.
    """
    fy_eps = [
        (npft / sh) if (npft not in (None, 0) and sh not in (None, 0)) else None
        for npft, sh in zip(fy_net_profit, fy_adj_shares)
    ]
    fy_frame = pd.DataFrame({
        "fy_end": pd.to_datetime(fy_dates),
        "eps": fy_eps,
    }).dropna(subset=["eps"]).sort_values("fy_end").reset_index(drop=True)
    fy_frame["available_date"] = fy_frame["fy_end"] + pd.Timedelta(days=reporting_lag_days)

    daily = eod_df[["Date", "Close"]].copy().sort_values("Date").reset_index(drop=True)
    eps_lookup = fy_frame[["available_date", "eps"]].rename(columns={"available_date": "Date"})
    merged = pd.merge_asof(daily, eps_lookup, on="Date", direction="backward")
    merged["current_pe"] = merged["Close"] / merged["eps"]
    return merged[["Date", "current_pe"]]


# Candidate NSE sectoral index DISPLAY names, cross-checked across NSE's own
# sectoral-indices page and independent index-data sources (web search,
# July 2026) - a reasonable starting point, NOT guaranteed to be the exact
# strings the nse package's API expects. The package's own docs reference a
# section titled "Acceptable values for nse.listEquityStocksByIndex" - check
# that yourself for the definitive, exact-format list before trusting this
# one. Indices also get added/restructured over time.
CANDIDATE_SECTORAL_INDICES = [
    "NIFTY AUTO", "NIFTY BANK", "NIFTY FINANCIAL SERVICES", "NIFTY FMCG",
    "NIFTY HEALTHCARE INDEX", "NIFTY IT", "NIFTY MEDIA", "NIFTY METAL",
    "NIFTY PHARMA", "NIFTY PRIVATE BANK", "NIFTY PSU BANK", "NIFTY REALTY",
    "NIFTY CONSUMER DURABLES", "NIFTY OIL AND GAS", "NIFTY CHEMICALS",
]


def build_sector_mapping(index_list: list, download_folder: str = "./nse_data") -> dict:
    """
    Builds symbol -> sector using NSE's own sectoral indices as the sector
    proxy - your explicit choice, revisiting the idea from earlier in this
    project. Uses nse.listEquityStocksByIndex() per index, then inverts.

    index_list is REQUIRED - no silent default. CANDIDATE_SECTORAL_INDICES
    above is a researched starting point (cross-checked across NSE's own
    page and other sources), but I have NOT verified these exact strings
    against the nse package's actual accepted values - check the package's
    own "Acceptable values for nse.listEquityStocksByIndex" docs section
    for the definitive list before relying on this one.

    UNTESTED - same network restriction as everything else touching NSE
    directly (promoter holding, order book) - this sandbox can't reach
    nseindia.com. Written correctly per the documented method signature,
    verify in your own environment first.

    A stock present in multiple sectoral indices will end up mapped to
    whichever index is processed LAST in index_list - if that matters to
    you, order index_list accordingly or handle multi-sector stocks
    explicitly, this doesn't resolve that for you.
    """
    from nse import NSE

    symbol_to_sector = {}
    failures = []

    with NSE(download_folder, server=True) as nse_client:
        for index_name in index_list:
            try:
                result = nse_client.listEquityStocksByIndex(index=index_name)
                for stock in result.get("data", []):
                    sym = stock.get("symbol")
                    if sym:
                        symbol_to_sector[sym] = index_name
            except Exception as e:
                failures.append((index_name, str(e)))

    return {"mapping": symbol_to_sector, "failures": failures}


def calc_valuation_score_V(panel_df: pd.DataFrame, lookback_days: int = 365,
                            min_population: int = 10) -> pd.DataFrame:
    """
    VADM_t's V-score, exactly per your spec: rolling 1-year PE percentile
    WITHIN SECTOR, inverted so cheaper = higher score: V = 1 - PE_percentile.

    Input panel_df must be LONG format, one row per stock per trading day:
    columns Date, Symbol, Sector, current_pe. Build this by combining
    calc_pointintime_pe_series() per stock (already point-in-time-safe)
    with your sector mapping from build_sector_mapping().

    For each (stock, date), the percentile is computed against the POOLED
    population of ALL that stock's sector-peers' PE values from the
    trailing `lookback_days` window - not just that single day's
    cross-section, which gives a much richer reference population,
    especially for sectors with few constituent stocks.

    min_population: dates where fewer than this many peer observations
    exist in the trailing window return NaN for V rather than a percentile
    computed off too few points - same discipline as min_history_points
    elsewhere in this file.

    Performance: loops over unique dates PER SECTOR, not every row, to
    stay tractable - still not something to run on every Streamlit
    page load, this is a model-fitting step, not a login-time calc.

    NOT YET TESTED against real multi-stock sector data - we don't have
    that yet (sector mapping untested against live NSE, no multi-stock
    Screener uploads done). Only verified against a synthetic dataset -
    see the accompanying test script - to confirm the mechanics are
    correct, NOT that real coefficients exist yet.
    """
    from scipy.stats import percentileofscore

    panel_df = panel_df.dropna(subset=["current_pe"]).sort_values("Date").reset_index(drop=True)
    panel_df["V"] = np.nan

    for sector, sector_df in panel_df.groupby("Sector"):
        sector_df = sector_df.sort_values("Date")
        unique_dates = sector_df["Date"].unique()

        for d in unique_dates:
            window_start = pd.Timestamp(d) - pd.Timedelta(days=lookback_days)
            window_mask = (sector_df["Date"] > window_start) & (sector_df["Date"] <= d)
            population = sector_df.loc[window_mask, "current_pe"].dropna()

            if len(population) < min_population:
                continue  # not enough sector-peer history yet - leave as NaN

            today_idx = sector_df.index[sector_df["Date"] == d]
            for idx in today_idx:
                pe_val = sector_df.loc[idx, "current_pe"]
                pctile = percentileofscore(population, pe_val) / 100
                panel_df.loc[idx, "V"] = 1 - pctile

    return panel_df


def fit_vadm_t_coefficients(panel_df: pd.DataFrame, forward_days: int = 10,
                             min_rows_per_sector: int = 200, test_fraction: float = 0.3) -> dict:
    """
    Sector-grouped OLS, exactly per your spec:
      X = V, D, V*D (interaction term)
      Y = forward_days-day forward return
    Fit SEPARATELY per sector - your explicit requirement - WITH the
    minimum-row-count safeguard flagged earlier: our pooled 29-stock
    delivery-only model needed 16,500 rows to even get near R²~0. A thin
    sector will produce Betas that look confident but are noise. Sectors
    below min_rows_per_sector are SKIPPED, not force-fit, and reported
    separately so you know which sectors genuinely lack enough data.

    Input panel_df needs: Date, Symbol, Sector, V, D, Close (Close is
    needed here to compute forward returns per stock).

    Chronological train/test split PER SECTOR - same no-lookahead
    discipline as run_alpha_white_backtest and fit_delivery_score_model_pooled,
    never a random shuffle.
    """
    import statsmodels.api as sm
    from sklearn.metrics import r2_score

    panel_df = panel_df.sort_values("Date").reset_index(drop=True)
    panel_df["VD_interaction"] = panel_df["V"] * panel_df["D"]

    results = {}
    skipped_sectors = []

    for sector, sector_df in panel_df.groupby("Sector"):
        sector_df = sector_df.sort_values("Date").copy()

        sector_df["forward_return"] = sector_df.groupby("Symbol")["Close"].transform(
            lambda s: (s.shift(-forward_days) - s) / s * 100
        )

        clean = sector_df.dropna(subset=["V", "D", "VD_interaction", "forward_return"])

        if len(clean) < min_rows_per_sector:
            skipped_sectors.append((sector, len(clean)))
            continue

        clean = clean.sort_values("Date")
        split_idx = int(len(clean) * (1 - test_fraction))
        split_date = clean.iloc[split_idx]["Date"]

        train = clean[clean["Date"] < split_date]
        test = clean[clean["Date"] >= split_date]
        feature_cols = ["V", "D", "VD_interaction"]

        X_train = sm.add_constant(train[feature_cols])
        model = sm.OLS(train["forward_return"], X_train).fit()

        X_test = sm.add_constant(test[feature_cols], has_constant="add")
        y_pred = model.predict(X_test)
        out_of_sample_r2 = r2_score(test["forward_return"], y_pred)

        results[sector] = {
            "weights": model.params.to_dict(),
            "p_values": model.pvalues.to_dict(),
            "in_sample_r2": float(model.rsquared),
            "out_of_sample_r2": float(out_of_sample_r2),
            "n_train": len(train),
            "n_test": len(test),
        }

    return {"results": results, "skipped_sectors": skipped_sectors}


def calc_vadm_t_score(V: float, D: float, weights: dict) -> float:
    """
    Applies FITTED Beta_1/Beta_2/Beta_3 (from fit_vadm_t_coefficients) to
    produce the actual VADM_t score, exactly per your formula:
        VADM_t = Beta_1*V + Beta_2*D + Beta_3*(V*D)
    weights dict keys must match fit_vadm_t_coefficients' output for one
    sector: const, V, D, VD_interaction.
    """
    return (
        weights.get("const", 0)
        + weights.get("V", 0) * V
        + weights.get("D", 0) * D
        + weights.get("VD_interaction", 0) * (V * D)
    )


# ---------------------------------------------------------------------------
# 10. CUSTOM SCORE - your 3-factor Altman-style model, separate from VADM_t.
#     Factors: Delivery %, Volume Growth (day-over-day), Price Return.
#     Weights sector-grouped-fitted per your explicit requirement
#     ("data k base pe AND sector k base pe"), not guessed.
# ---------------------------------------------------------------------------

def calc_volume_growth(df: pd.DataFrame) -> pd.Series:
    """
    Volume Growth, per your Step 1 spec: "Pichle din ke mukable volume
    kitna bada" - day-over-day % change in volume.
    """
    return df["Volume"].pct_change() * 100


def calc_price_return(df: pd.DataFrame) -> pd.Series:
    """Price Return, per your Step 1 spec: daily % return."""
    return df["Close"].pct_change() * 100


def build_custom_score_factors(df: pd.DataFrame, lookback: int = 20) -> pd.DataFrame:
    """
    Your Step 2 exactly: all 3 factors converted to rolling Z-scores so
    they're on the same scale before combining - Delivery% is 0-100,
    Volume is in lakhs, Price Return is a small %, none directly comparable
    without this.
    """
    dlv_pct = calc_delivery_pct(df)
    vol_growth = calc_volume_growth(df)
    price_return = calc_price_return(df)

    def _zscore(s):
        roll_mean = s.rolling(lookback).mean()
        roll_std = s.rolling(lookback).std()
        return (s - roll_mean) / roll_std.replace(0, np.nan)

    return pd.DataFrame({
        "delivery_pct_zscore": _zscore(dlv_pct),
        "volume_growth_zscore": _zscore(vol_growth),
        "price_return_zscore": _zscore(price_return),
    })


def build_custom_score_panel(stock_eod_dfs: dict, sector_mapping: dict, lookback: int = 20):
    """
    Builds the long-format panel needed to fit the Custom Score across
    multiple stocks/sectors - combines each stock's EOD2 data with the
    sector mapping from build_sector_mapping()['mapping'].

    stock_eod_dfs: {symbol: eod_df} for each stock to include.
    sector_mapping: {symbol: sector}. Stocks missing from the mapping are
    skipped and returned separately, not silently dropped.

    Returns (panel_df, skipped_symbols).
    """
    rows = []
    skipped = []

    for symbol, df in stock_eod_dfs.items():
        sector = sector_mapping.get(symbol)
        if sector is None:
            skipped.append(symbol)
            continue
        factors = build_custom_score_factors(df, lookback=lookback)
        piece = factors.copy()
        piece["Date"] = df["Date"].values
        piece["Close"] = df["Close"].values
        piece["Symbol"] = symbol
        piece["Sector"] = sector
        rows.append(piece)

    if not rows:
        return pd.DataFrame(), skipped

    return pd.concat(rows, ignore_index=True), skipped


def fit_custom_score_coefficients(panel_df: pd.DataFrame, forward_days: int = 10,
                                   min_rows_per_sector: int = 200, test_fraction: float = 0.3) -> dict:
    """
    Fits the Custom Score's weights via SECTOR-GROUPED OLS - per your
    explicit requirement, data-driven and sector-based, not guessed.

    X = delivery_pct_zscore, volume_growth_zscore, price_return_zscore
    Y = forward_days-day forward return

    Exact same discipline as fit_vadm_t_coefficients: chronological
    train/test split PER SECTOR (never random shuffling), sectors below
    min_rows_per_sector are SKIPPED and reported separately rather than
    force-fit into an unreliable result, out-of-sample R² reported as the
    number to actually trust.

    panel_df needs: Date, Symbol, Sector, delivery_pct_zscore,
    volume_growth_zscore, price_return_zscore, Close.
    """
    import statsmodels.api as sm
    from sklearn.metrics import r2_score

    panel_df = panel_df.sort_values("Date").reset_index(drop=True)
    feature_cols = ["delivery_pct_zscore", "volume_growth_zscore", "price_return_zscore"]

    results = {}
    skipped_sectors = []

    for sector, sector_df in panel_df.groupby("Sector"):
        sector_df = sector_df.sort_values("Date").copy()
        sector_df["forward_return"] = sector_df.groupby("Symbol")["Close"].transform(
            lambda s: (s.shift(-forward_days) - s) / s * 100
        )

        clean = sector_df.dropna(subset=feature_cols + ["forward_return"])

        if len(clean) < min_rows_per_sector:
            skipped_sectors.append((sector, len(clean)))
            continue

        clean = clean.sort_values("Date")
        split_idx = int(len(clean) * (1 - test_fraction))
        split_date = clean.iloc[split_idx]["Date"]

        train = clean[clean["Date"] < split_date]
        test = clean[clean["Date"] >= split_date]

        X_train = sm.add_constant(train[feature_cols])
        model = sm.OLS(train["forward_return"], X_train).fit()

        X_test = sm.add_constant(test[feature_cols], has_constant="add")
        y_pred = model.predict(X_test)
        out_of_sample_r2 = r2_score(test["forward_return"], y_pred)

        results[sector] = {
            "weights": model.params.to_dict(),
            "p_values": model.pvalues.to_dict(),
            "in_sample_r2": float(model.rsquared),
            "out_of_sample_r2": float(out_of_sample_r2),
            "n_train": len(train),
            "n_test": len(test),
        }

    return {"results": results, "skipped_sectors": skipped_sectors}


def calc_custom_score(delivery_pct_zscore: float, volume_growth_zscore: float,
                       price_return_zscore: float, weights: dict) -> float:
    """
    Applies FITTED weights (from fit_custom_score_coefficients, per sector)
    to produce the actual Custom Score - your Step 3, but with real
    data-driven coefficients instead of guessed ones.
    """
    return (
        weights.get("const", 0)
        + weights.get("delivery_pct_zscore", 0) * delivery_pct_zscore
        + weights.get("volume_growth_zscore", 0) * volume_growth_zscore
        + weights.get("price_return_zscore", 0) * price_return_zscore
    )


# ---------------------------------------------------------------------------
# 11. VADM_t UNIVERSAL - works for ANY single stock (not hardcoded to one),
#     using self-relative V as a working fallback since true sector-relative
#     V needs live sector mapping + peer data that isn't set up yet.
#     Quadrant classification + empirical hypothesis testing (H1-H4).
# ---------------------------------------------------------------------------

def calc_valuation_score_V_selfrelative(eod_df: pd.DataFrame, fy_dates: list, fy_price: list,
                                         fy_net_profit: list, fy_adj_shares: list,
                                         lookback: int = 252, reporting_lag_days: int = 60,
                                         min_periods: int = 30) -> pd.DataFrame:
    """
    UNIVERSAL, single-stock V-score - works for whichever stock is passed
    in, not hardcoded. This is a SELF-RELATIVE fallback for VADM_t's V:
    true sector-relative V (calc_valuation_score_V, vs sector peers) needs
    live sector mapping + multiple stocks' EPS data that isn't set up yet.
    Until that's ready, this reuses the same point-in-time-safe PE series
    and applies a rolling percentile rank against the stock's OWN history -
    same self-relative logic Alpha White's PE regime already uses,
    inverted per your V = 1 - percentile formula.

    Returns DataFrame: Date, current_pe, V (0 to 1, higher = cheaper
    relative to its own trailing history).
    """
    pe_df = calc_pointintime_pe_series(eod_df, fy_dates, fy_price, fy_net_profit,
                                        fy_adj_shares, reporting_lag_days)
    pe_df["V"] = 1 - pe_df["current_pe"].rolling(lookback, min_periods=min_periods).rank(pct=True)
    return pe_df[["Date", "current_pe", "V"]]


def calc_quadrant(V: float, D: float) -> str:
    """
    Classifies into VADM_t's 2x2 quadrant framework:
      Cheap        = V > 0.5 (better than its own median self-relative valuation)
      Expensive    = V <= 0.5
      Accumulation = D > 0 (delivery flow above its own recent average)
      Distribution = D <= 0
    """
    if pd.isna(V) or pd.isna(D):
        return "INSUFFICIENT_DATA"
    valuation = "Cheap" if V > 0.5 else "Expensive"
    flow = "Accumulation" if D > 0 else "Distribution"
    return f"{valuation}_{flow}"


def build_quadrant_markers(detail_df: pd.DataFrame) -> list:
    """
    Converts quadrant classification into chart markers, for the "signals
    on the chart" you asked for. Cheap+Accumulation (the quadrant H1+H2+H3
    predict is best) gets a buy-style marker; Expensive+Distribution
    (H4's predicted worst) gets a sell-style marker. The two MIXED
    quadrants (Cheap+Distribution, Expensive+Accumulation - where V and D
    disagree) get no marker deliberately - those are exactly the ambiguous
    cases H3's interaction term exists to explain, not places confident
    enough to mark either way.
    """
    markers = []
    for _, row in detail_df.iterrows():
        date_str = row["Date"].strftime("%Y-%m-%d") if hasattr(row["Date"], "strftime") else str(row["Date"])
        if row["quadrant"] == "Cheap_Accumulation":
            markers.append({"time": date_str, "position": "belowBar", "color": "#26a69a",
                             "shape": "arrowUp", "text": "C+A"})
        elif row["quadrant"] == "Expensive_Distribution":
            markers.append({"time": date_str, "position": "aboveBar", "color": "#ef5350",
                             "shape": "arrowDown", "text": "E+D"})
    return markers


def test_vadm_t_hypotheses(eod_df: pd.DataFrame, fy_dates: list, fy_price: list,
                            fy_net_profit: list, fy_adj_shares: list,
                            delivery_lookback: int = 20, valuation_lookback: int = 252,
                            forward_days: int = 20, reporting_lag_days: int = 60) -> dict:
    """
    Empirically tests H1-H4 for ONE stock - UNIVERSAL, works for whichever
    stock's data is passed in. Uses self-relative V (see caveat above) and
    the simple-average D (per the robustness decision - D is NOT pre-fit
    against forward returns here, avoiding the double-fit risk you asked
    about).

      H1: Cheap (V>0.5) should show higher forward returns than Expensive
      H2: Accumulation (D>0) should show higher forward returns than Distribution
      H3: Interaction - Cheap+Accumulation should beat what H1 and H2 would
          predict independently (tests whether combining adds value beyond
          either alone)
      H4 (proposed by me, your call to confirm/change): Expensive+Distribution
          should show the LOWEST forward returns, completing the quadrant's
          symmetric logic

    Returns per-quadrant average forward return + sample count, and a
    plain verdict on whether each hypothesis holds for THIS stock's
    history - real result, not asserted as universally true.
    """
    v_df = calc_valuation_score_V_selfrelative(
        eod_df, fy_dates, fy_price, fy_net_profit, fy_adj_shares,
        lookback=valuation_lookback, reporting_lag_days=reporting_lag_days
    )
    D_series = calc_delivery_score_D(eod_df, lookback=delivery_lookback)

    merged = eod_df[["Date", "Close"]].copy().reset_index(drop=True)
    merged = merged.merge(v_df[["Date", "V"]], on="Date", how="left")
    merged["D"] = D_series.reset_index(drop=True).values
    merged["quadrant"] = [calc_quadrant(v, d) for v, d in zip(merged["V"], merged["D"])]

    close = merged["Close"]
    merged["forward_return"] = (close.shift(-forward_days) - close) / close * 100

    clean = merged.dropna(subset=["V", "D", "forward_return"])
    clean = clean[clean["quadrant"] != "INSUFFICIENT_DATA"]

    if clean.empty:
        return {"error": "No clean rows to test hypotheses on - needs more history/EPS data.",
                "quadrant_stats": {}, "verdicts": {}, "detail_df": pd.DataFrame()}

    quadrant_stats = clean.groupby("quadrant")["forward_return"].agg(["mean", "count"]).to_dict("index")

    def _get_mean(q):
        return quadrant_stats.get(q, {}).get("mean")

    cheap_mean = clean[clean["V"] > 0.5]["forward_return"].mean()
    expensive_mean = clean[clean["V"] <= 0.5]["forward_return"].mean()
    accum_mean = clean[clean["D"] > 0]["forward_return"].mean()
    dist_mean = clean[clean["D"] <= 0]["forward_return"].mean()

    cheap_accum = _get_mean("Cheap_Accumulation")
    expensive_dist = _get_mean("Expensive_Distribution")

    verdicts = {
        "H1 (Cheap > Expensive)": "HOLDS" if (cheap_mean is not None and expensive_mean is not None
                                               and cheap_mean > expensive_mean) else "DOES NOT HOLD",
        "H2 (Accumulation > Distribution)": "HOLDS" if (accum_mean is not None and dist_mean is not None
                                                          and accum_mean > dist_mean) else "DOES NOT HOLD",
        "H3 (Cheap+Accum beats independent effects)": (
            "HOLDS" if (cheap_accum is not None and cheap_mean is not None and accum_mean is not None
                        and cheap_accum > max(cheap_mean, accum_mean)) else "DOES NOT HOLD"
        ),
        "H4 (Expensive+Distribution is worst quadrant)": (
            "HOLDS" if (expensive_dist is not None and quadrant_stats
                        and expensive_dist == min(v["mean"] for v in quadrant_stats.values())) else "DOES NOT HOLD"
        ),
    }

    return {
        "quadrant_stats": quadrant_stats,
        "verdicts": verdicts,
        "cheap_mean": cheap_mean, "expensive_mean": expensive_mean,
        "accum_mean": accum_mean, "dist_mean": dist_mean,
        "n_total": len(clean),
        "detail_df": clean,
    }


# ---------------------------------------------------------------------------
# STILL OPEN - NOT BUILT, NOT GUESSED:
#
# 1. No REAL coefficients exist yet. The full pipeline (V-score, D-score,
#    sector-grouped regression) is now code-complete and mechanically
#    tested against synthetic data, but has never run against real
#    multi-stock sector data - that needs: (a) sector mapping actually run
#    against live NSE (untested from this sandbox), (b) multiple stocks'
#    EPS data per sector (Screener batch upload, your chosen source).
#    Until both of those happen, calling fit_vadm_t_coefficients on real
#    data hasn't been done - don't assume any Beta values exist.
# ---------------------------------------------------------------------------
