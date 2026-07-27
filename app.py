"""
VADM - Streamlit app
=====================
Built from vadm_calculations.py. UI layer only - all math lives in the
calc module, per our calc-layer/UI-layer split.

WHAT WORKS END TO END: EOD2 fetch, delivery %, chart with toggleable
indicators, screener Excel parsing + fundamentals + PE/EV per year.

WHAT'S INTENTIONALLY INCOMPLETE (see inline notes, not hidden):
  - sector_avg_pe is a manual number input for now - no resolved automatic
    source yet.
  - Promoter holding fetch is wired in but wrapped in try/except since it's
    never been executed against the live NSE site from this environment.
  - Alpha Black tab is a stub - no formula/H4 to build against yet.
  - "Liquidity entered" metric is explicitly labeled as an assumption.
"""

import numpy as np
import pandas as pd
import streamlit as st

from vadm_calculations import (
    fetch_stock_universe,
    load_eod2_data,
    calc_delivery_pct,
    calc_relative_delivery,
    calc_volume_percentile,
    calc_52wk_high_low,
    fetch_market_depth,
    estimate_exit_price,
    parse_screener_excel,
    calc_pe_per_year,
    calc_market_cap_per_year,
    calc_ev_per_year,
    calc_revenue_growth,
    fetch_promoter_holding,
    alpha_white_signal,
)

from lightweight_charts_v5 import lightweight_charts_v5_component
from indicators import PriceIndicator, VolumeIndicator, RSIIndicator, MACDIndicator, SMAIndicator


st.set_page_config(page_title="VADM", layout="wide")


# ---------------------------------------------------------------------------
# SEARCH - one search, up top, feeds BOTH strategy tabs below.
# Not in the sidebar anymore - this is the main flow: search once, then pick
# a strategy tab. Both tabs read the same `symbol` / `eod_df` set here.
# ---------------------------------------------------------------------------

st.title("VADM")

@st.cache_data(ttl=3600)
def get_universe():
    return fetch_stock_universe()

universe = get_universe()

# Search now only fires on the button click, not on every dropdown change -
# wrapped in st.form so picking a symbol doesn't trigger anything by itself.
# The chosen symbol/excel are pushed into session_state on submit so that
# later widget interactions (checkboxes, sliders inside the tabs) don't
# clear the results - only a fresh "Search" click changes what's loaded.
if "vadm_symbol" not in st.session_state:
    st.session_state.vadm_symbol = None
if "vadm_excel_bytes" not in st.session_state:
    st.session_state.vadm_excel_bytes = None

with st.form("search_form"):
    search_col, upload_col, btn_col = st.columns([2, 1, 0.6])

    with search_col:
        symbol_choice = st.selectbox(
            "🔍 Search stock",
            options=sorted(universe.keys()),
            index=sorted(universe.keys()).index("RELIANCE") if "RELIANCE" in universe else 0,
            help="Pick a symbol, then click Search. This one search feeds both tabs below.",
        )

    with upload_col:
        excel_choice = st.file_uploader(
            "Screener.in export (.xlsx)", type=["xlsx"],
            help="Single-company export, same structure as IRB_Infra_Devl-3.xlsx"
        )

    with btn_col:
        st.write("")  # vertical spacer to align button with the inputs
        st.write("")
        submitted = st.form_submit_button("Search", use_container_width=True)

if submitted:
    st.session_state.vadm_symbol = symbol_choice
    # Store raw bytes, not the UploadedFile object itself - the object's
    # read pointer doesn't survive being reused across reruns cleanly.
    st.session_state.vadm_excel_bytes = excel_choice.getvalue() if excel_choice else None

symbol = st.session_state.vadm_symbol
uploaded_excel_bytes = st.session_state.vadm_excel_bytes

if symbol is None:
    st.info("👆 Search for a stock above to see Alpha White / Alpha Black results.")
    st.stop()


@st.cache_data(ttl=900)
def get_eod2(sym):
    return load_eod2_data(sym)

try:
    eod_df = get_eod2(symbol)
    data_load_error = None
except Exception as e:
    eod_df = None
    data_load_error = str(e)

if data_load_error:
    st.error(f"Could not load EOD2 data for {symbol}: {data_load_error}")

st.divider()

# Tabs sit BELOW the search, per your layout - one search, then pick strategy.
tab_white, tab_black = st.tabs(["Alpha White", "Alpha Black"])

# ---------------------------------------------------------------------------
# TAB: ALPHA WHITE
# ---------------------------------------------------------------------------
with tab_white:
    st.header("Alpha White")

    if eod_df is None:
        st.warning("No price data loaded - pick a valid symbol.")
    else:
        # --- Top row: CMP / 52w H-L ---
        col1, col2, col3 = st.columns(3)
        cmp = eod_df["Close"].iloc[-1]
        hi52, lo52 = calc_52wk_high_low(eod_df)

        col1.metric("CMP", f"₹{cmp:,.2f}")
        col2.metric("52w High", f"₹{hi52:,.2f}")
        col3.metric("52w Low", f"₹{lo52:,.2f}")

        st.divider()

        # --- Exit Liquidity: live order book, not historical volume ---
        # Replaces the old 6/12-month liquidity guess entirely, per your
        # instruction. This is a live call, so it only fires on the button
        # click - not on every rerun caused by other widgets on this page.
        st.subheader("Exit Liquidity — Live Order Book")
        st.caption(
            "Shows where bid-side size actually sits right now, and estimates "
            "your fill price if you exit into it. Untested against the live "
            "NSE site from my end (sandboxed) - the field-name mapping is "
            "confirmed against the package's own sample response, but not a "
            "live call. Run it once yourself to be sure."
        )

        depth_col, qty_col = st.columns([1, 1])
        with qty_col:
            exit_qty = st.number_input("Shares to exit", min_value=1, value=100, step=1)
        with depth_col:
            fetch_depth_clicked = st.button("Fetch live order book")

        if fetch_depth_clicked:
            try:
                depth = fetch_market_depth(symbol)
                bid_df = pd.DataFrame(depth["bids"])
                ask_df = pd.DataFrame(depth["asks"])

                bcol, acol = st.columns(2)
                with bcol:
                    st.markdown("**Bids (buyers)** — this is the side you sell into")
                    st.dataframe(bid_df, use_container_width=True, hide_index=True)
                with acol:
                    st.markdown("**Asks (sellers)**")
                    st.dataframe(ask_df, use_container_width=True, hide_index=True)

                st.metric("Last traded price", f"₹{depth['last_price']:,.2f}" if depth['last_price'] else "N/A")

                result = estimate_exit_price(depth["bids"], exit_qty)
                if result["depth_sufficient"]:
                    st.success(
                        f"Estimated exit VWAP for {exit_qty} shares: "
                        f"₹{result['estimated_vwap_price']:,.2f} "
                        f"(fully absorbed within visible depth)"
                    )
                else:
                    st.warning(
                        f"Only {result['filled_qty']} of {exit_qty} shares fillable within "
                        f"the visible 5-level depth (₹{result['estimated_vwap_price']:,.2f} VWAP "
                        f"for that portion). Remaining {result['unfilled_qty']} shares would "
                        f"go beyond what NSE's free quote shows - real fill price for those "
                        f"would likely be worse."
                    )
            except Exception as e:
                st.error(
                    f"Live fetch failed: {e}. This code path has never been executed "
                    f"against the real NSE site from my sandbox - if this is a field-name "
                    f"mismatch rather than a network issue, tell me the actual error and "
                    f"I'll fix the mapping in fetch_market_depth()."
                )

        st.divider()

        # --- Fundamentals (only if Excel uploaded) ---
        st.subheader("Fundamentals")
        if uploaded_excel_bytes is not None:
            try:
                with open("/tmp/_uploaded_screener.xlsx", "wb") as f:
                    f.write(uploaded_excel_bytes)
                parsed = parse_screener_excel("/tmp/_uploaded_screener.xlsx")

                years = [d.year if hasattr(d, "year") else d for d in parsed["years"]]
                pe_series = calc_pe_per_year(parsed["price"], parsed["net_profit"], parsed["adjusted_shares_cr"])
                mcap_series = calc_market_cap_per_year(parsed["price"], parsed["adjusted_shares_cr"])
                ev_series = calc_ev_per_year(parsed["price"], parsed["adjusted_shares_cr"],
                                              parsed["borrowings"], parsed["cash_and_bank"])
                growth_series = calc_revenue_growth(parsed["sales"])

                # --- EV snapshot, in cards, per your request - latest FY's
                # EV and the components that build up to it (Market Cap,
                # Total Debt, Cash & Bank), same visual style as the
                # CMP/52w cards up top. Full year-by-year trend still sits
                # in the table below this - cards are the quick-glance view.
                latest_mcap = mcap_series[-1] if mcap_series and mcap_series[-1] is not None else None
                latest_debt = parsed["borrowings"][-1] if parsed["borrowings"] else None
                latest_cash = parsed["cash_and_bank"][-1] if parsed["cash_and_bank"] else None
                latest_ev = ev_series[-1] if ev_series and ev_series[-1] is not None else None
                latest_fy = years[-1] if years else "latest"

                st.markdown(f"**EV Snapshot — FY{latest_fy}**")
                ev_c1, ev_c2, ev_c3, ev_c4 = st.columns(4)
                ev_c1.metric("Market Cap (Cr)", f"₹{latest_mcap:,.0f}" if latest_mcap is not None else "N/A")
                ev_c2.metric("Total Debt (Cr)", f"₹{latest_debt:,.0f}" if latest_debt is not None else "N/A")
                ev_c3.metric("Cash & Bank (Cr)", f"₹{latest_cash:,.0f}" if latest_cash is not None else "N/A")
                ev_c4.metric("Enterprise Value (Cr)", f"₹{latest_ev:,.0f}" if latest_ev is not None else "N/A")
                st.caption("EV = Market Cap + Total Debt − Cash & Bank (simplified, no minority interest adjustment)")

                fundamentals_df = pd.DataFrame({
                    "FY": years,
                    "Sales (Cr)": parsed["sales"],
                    "Revenue Growth %": [round(g, 2) if g is not None else None for g in growth_series],
                    "Net Profit (Cr)": parsed["net_profit"],
                    "Total Debt / Borrowings (Cr)": parsed["borrowings"],
                    "Cash & Bank (Cr)": parsed["cash_and_bank"],
                    "PE": [round(p, 2) if p is not None else None for p in pe_series],
                    "EV (Cr)": [round(e, 2) if e is not None else None for e in ev_series],
                })
                st.dataframe(fundamentals_df, use_container_width=True)

                current_pe = pe_series[-1] if pe_series and pe_series[-1] is not None else None

                # Promoter holding - untested against live NSE from this sandbox
                st.markdown("**Promoter Holding**")
                try:
                    holding = fetch_promoter_holding(symbol)
                    st.write(f"{holding['promoter_holding_pct']}% (as of {holding['as_of']})")
                except Exception as e:
                    st.info(
                        f"Promoter holding fetch not available here: {e}. "
                        f"This code is untested against live NSE - run it in your own "
                        f"environment (this sandbox can't reach nseindia.com)."
                    )

            except Exception as e:
                st.error(f"Couldn't parse the uploaded Excel: {e}")
                current_pe = None
        else:
            st.info("Upload a Screener.in export to see fundamentals + PE history.")
            current_pe = None

        st.divider()

        # --- Alpha White signal ---
        st.subheader("Signal")
        sig_col1, sig_col2, sig_col3 = st.columns(3)

        stock_pe_input = sig_col1.number_input(
            "Stock PE", value=float(current_pe) if current_pe else 0.0, step=0.1
        )
        sector_avg_pe_input = sig_col2.number_input(
            "Sector avg PE (manual entry - no auto source yet)", value=0.0, step=0.1
        )
        volume_lookback = sig_col3.slider("Volume percentile lookback (days)", 20, 252, 60)
        volume_pctile_threshold = st.slider(
            "Volume percentile threshold (starting hypothesis - backtest to tune)", 50, 99, 80
        )

        current_vol_pctile = calc_volume_percentile(eod_df, lookback=volume_lookback).iloc[-1]

        signal = alpha_white_signal(
            stock_pe=stock_pe_input if stock_pe_input else None,
            sector_avg_pe=sector_avg_pe_input if sector_avg_pe_input else None,
            current_volume_percentile=current_vol_pctile,
            volume_percentile_threshold=volume_pctile_threshold,
        )

        if signal == "BUY":
            st.success(f"Signal: **BUY** (current volume percentile: {current_vol_pctile:.1f})")
        elif signal == "SELL":
            st.error(f"Signal: **SELL** (current volume percentile: {current_vol_pctile:.1f})")
        else:
            st.warning("Signal: insufficient data - enter Stock PE and Sector avg PE above.")

        st.divider()

        # --- Chart ---
        st.subheader("Chart")

        chart_df = eod_df.rename(columns={
            "Date": "date", "Open": "open", "High": "high",
            "Low": "low", "Close": "close", "Volume": "volume",
        }).copy()
        chart_df["date"] = chart_df["date"].dt.strftime("%Y-%m-%d")
        chart_df = chart_df.replace({np.nan: None})
        chart_df["delivery_pct"] = calc_delivery_pct(eod_df).values

        ic1, ic2, ic3, ic4 = st.columns(4)
        show_sma = ic1.checkbox("SMA 20/50/200", value=True)
        show_vol_dlv = ic2.checkbox("Volume + Delivery %", value=True)
        show_rsi = ic3.checkbox("RSI", value=False)
        show_macd = ic4.checkbox("MACD", value=False)

        overlays = []
        if show_sma:
            overlays = [
                SMAIndicator(chart_df, period=20, color="rgba(255,140,0,0.8)"),
                SMAIndicator(chart_df, period=50, color="rgba(25,118,210,0.8)"),
                SMAIndicator(chart_df, period=200, color="rgba(156,39,176,0.8)"),
            ]

        price_indicator = PriceIndicator(
            chart_df, height=450, title=f"{symbol}", style="Candlestick", overlays=overlays,
        )
        price_indicator.calculate()
        indicators = [price_indicator]

        if show_vol_dlv:
            vol_indicator = VolumeIndicator(chart_df, height=140)
            vol_indicator.calculate()
            vol_pane = vol_indicator.pane()
            # FIX: earlier version sent raw NaN (from the rolling calc's
            # warm-up period and any zero-volume days) straight into the
            # JSON payload. Python's json module writes bare `NaN`, which
            # is NOT valid JSON - the browser's strict JSON.parse rejects
            # it with exactly the error you saw. The library's own
            # SMAIndicator already filters this with `.notnull()` (checked
            # its source) - my custom series never did. Fixed the same way.
            dlv_data = chart_df[["date", "delivery_pct"]].rename(
                columns={"date": "time", "delivery_pct": "value"}
            )
            dlv_data = dlv_data[dlv_data["value"].notnull()]
            dlv_records = dlv_data.to_dict(orient="records")

            vol_pane["series"].append({
                "type": "Line",
                "data": dlv_records,
                "options": {"color": "#FFD700", "lineWidth": 2},
                "priceScale": {
                    "scaleMargins": {"top": 0.1, "bottom": 0.6},
                    "priceScaleId": "delivery_pct",
                },
            })
            indicators.append(vol_pane)  # already a pane dict, not an Indicator

        if show_rsi:
            rsi_indicator = RSIIndicator(chart_df, height=120)
            rsi_indicator.calculate()
            indicators.append(rsi_indicator)

        if show_macd:
            macd_indicator = MACDIndicator(chart_df, height=120)
            macd_indicator.calculate()
            indicators.append(macd_indicator)

        charts_config = [ind if isinstance(ind, dict) else ind.pane() for ind in indicators]
        total_height = sum(cfg.get("height", 200) for cfg in charts_config)

        @st.fragment
        def render_chart():
            return lightweight_charts_v5_component(
                name=symbol,
                charts=charts_config,
                height=total_height,
                zoom_level=150,
                key=f"chart_{symbol}",
            )

        render_chart()


# ---------------------------------------------------------------------------
# TAB: ALPHA BLACK - stub only. No formula/H4 to build against yet.
# ---------------------------------------------------------------------------
with tab_black:
    st.header("Alpha Black")
    st.warning(
        "Not built yet, on purpose. VADM_t's functional form f() and H4 have "
        "never been specified - your own project memory flags this as a "
        "CEO-level decision. What IS ready below: the delivery-flow data "
        "H2 depends on, computed and displayed, waiting for the formula "
        "that will consume it."
    )
    if eod_df is not None:
        dlv_pct = calc_delivery_pct(eod_df)
        rel_dlv = calc_relative_delivery(eod_df, lookback=60, method="zscore")
        st.line_chart(pd.DataFrame({
            "Delivery %": dlv_pct.tail(120).values,
            "Relative Delivery (z-score)": rel_dlv.tail(120).values,
        }, index=eod_df["Date"].tail(120)))
