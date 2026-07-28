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
# STILL OPEN - NOT BUILT, NOT GUESSED:
#
# 1. Alpha Black / VADM_t. No formula f(), no H4. Your own project memory
#    flags this explicitly as a CEO-level decision - nothing here invents it.
#
# (Sector-relative PE is no longer blocking anything - Alpha White's PE
# condition was switched to self-relative, per your latest instruction,
# instead of waiting on a sector data source that was never resolved.)
# ---------------------------------------------------------------------------
