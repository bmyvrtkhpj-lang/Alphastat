"""
VADM - Streamlit app
=====================
Built from vadm_calculations.py. UI layer only - all math lives in the
calc module, per our calc-layer/UI-layer split.

WHAT WORKS END TO END: EOD2 fetch, delivery %, chart with toggleable
indicators, screener Excel parsing + fundamentals + PE/EV per year,
self-relative PE regime signal, live order-book exit liquidity.

WHAT'S INTENTIONALLY INCOMPLETE (see inline notes, not hidden):
  - Promoter holding fetch is wired in but wrapped in try/except since it's
    never been executed against the live NSE site from this environment.
  - Exit-liquidity order book fetch: same - untested against live NSE.
  - Alpha Black tab is a stub - no formula/H4 to build against yet.
  - PE regime is self-relative (vs own ~10 annual data points), not
    sector-relative - coarse by nature, flagged in calc_relative_pe_regime().
"""

import numpy as np
import pandas as pd
import streamlit as st

from vadm_calculations import (
    fetch_stock_universe,
    check_symbol_status,
    load_eod2_data,
    calc_delivery_pct,
    calc_relative_delivery,
    calc_volume_regime,
    calc_52wk_high_low,
    fetch_market_depth,
    estimate_exit_price,
    parse_screener_excel,
    calc_pe_per_year,
    calc_current_pe,
    calc_relative_pe_regime,
    calc_market_cap_per_year,
    calc_ev_per_year,
    calc_revenue_growth,
    fetch_promoter_holding,
    alpha_white_signal,
    run_alpha_white_backtest,
    summarize_backtest_signals,
    build_signal_markers,
    build_sector_mapping,
    CANDIDATE_SECTORAL_INDICES,
    calc_delivery_score_D,
    calc_valuation_score_V_selfrelative,
    calc_quadrant,
    test_vadm_t_hypotheses,
    build_quadrant_markers,
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
    # Symbol exists in the universe but has no current daily file - usually
    # means it's delisted/merged/renamed. Check its own history record
    # instead of leaving just the raw HTTP error.
    try:
        status = check_symbol_status(symbol)
        if status["found"]:
            st.info(status["message"])
    except Exception:
        pass  # best-effort explanation only - don't let this itself crash the app

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
                    "PE (FY-end)": [round(p, 2) if p is not None else None for p in pe_series],
                    "EV (Cr)": [round(e, 2) if e is not None else None for e in ev_series],
                })
                st.dataframe(fundamentals_df, use_container_width=True)

                # LIVE current PE - today's CMP over latest annual EPS, NOT
                # the same as the FY-end PE column above (that uses the
                # stale FY-end price). This is what feeds the signal below.
                live_current_pe = calc_current_pe(
                    cmp, parsed["net_profit"][-1], parsed["adjusted_shares_cr"][-1]
                )

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
                pe_series, live_current_pe = [], None
        else:
            st.info("Upload a Screener.in export to see fundamentals + PE history.")
            pe_series, live_current_pe = [], None

        st.divider()

        # --- Alpha White signal ---
        st.subheader("Signal")

        st.metric("Current PE (live CMP ÷ latest annual EPS)",
                   f"{live_current_pe:.2f}" if live_current_pe is not None else "N/A - upload Excel")

        sig_col1, sig_col2, sig_col3 = st.columns(3)

        pe_low_pctile = sig_col1.slider(
            "PE 'low' percentile (relative to own history)", 1, 49, 20
        )
        pe_high_pctile = sig_col2.slider(
            "PE 'high' percentile (relative to own history)", 51, 99, 80
        )
        volume_lookback = sig_col3.slider("Volume percentile lookback (days)", 20, 252, 60)

        volume_pctile_threshold = st.slider(
            "Volume percentile threshold (starting hypothesis - backtest to tune)", 50, 99, 80
        )

        pe_regime_result = calc_relative_pe_regime(
            live_current_pe, pe_series, low_percentile=pe_low_pctile, high_percentile=pe_high_pctile
        )
        volume_regime_result = calc_volume_regime(
            eod_df, lookback=volume_lookback, percentile_threshold=volume_pctile_threshold
        )

        if pe_regime_result["regime"] != "INSUFFICIENT_DATA":
            st.caption(
                f"PE regime: **{pe_regime_result['regime']}** "
                f"(low threshold ₹{pe_regime_result['low_threshold']:.1f}, "
                f"high threshold ₹{pe_regime_result['high_threshold']:.1f}, "
                f"current ₹{live_current_pe:.1f}) — based on only ~{len(pe_series)} annual "
                f"data points, coarse by nature."
            )
        if volume_regime_result["regime"] != "INSUFFICIENT_DATA":
            st.caption(
                f"Volume regime: **{volume_regime_result['regime']}** "
                f"(percentile {volume_regime_result['volume_percentile']:.1f}, "
                f"CLV {volume_regime_result['clv']:+.2f} — positive means closed "
                f"nearer the day's high, negative nearer the day's low)"
            )

        signal = alpha_white_signal(
            pe_regime=pe_regime_result["regime"],
            volume_regime=volume_regime_result["regime"],
        )

        if signal == "BUY":
            st.success("Signal: **BUY** (PE cheap + heavy buy-side volume)")
        elif signal == "SELL":
            st.error("Signal: **SELL** (PE expensive + heavy sell-side volume)")
        elif signal == "HOLD":
            st.info(
                f"Signal: **HOLD** - PE regime is {pe_regime_result['regime']}, "
                f"Volume regime is {volume_regime_result['regime']}; conditions "
                f"for BUY or SELL aren't both met."
            )
        else:
            st.warning("Signal: insufficient data - upload the Screener Excel above to compute PE regime.")

        st.divider()

        # --- Backtest ---
        st.subheader("Backtest")
        st.caption(
            "Runs the same signal logic across full price history - zero lookahead "
            "bias: PE percentile thresholds only ever use FY data that would have "
            "actually been known/reported as of each historical date, never future "
            "data. See run_alpha_white_backtest()'s docstring for the exact mechanism."
        )

        if "vadm_backtest" not in st.session_state:
            st.session_state.vadm_backtest = None

        holding_days = st.slider("Forward-return holding period (trading days)", 5, 60, 20)

        if st.button("Run Backtest"):
            if not pe_series:
                st.warning("Upload the Screener Excel above first - backtest needs the annual PE history.")
            else:
                bt_df = run_alpha_white_backtest(
                    eod_df, parsed["years"], parsed["price"], parsed["net_profit"], parsed["adjusted_shares_cr"],
                    volume_lookback=volume_lookback, volume_pctile_threshold=volume_pctile_threshold,
                    pe_low_pctile=pe_low_pctile, pe_high_pctile=pe_high_pctile,
                )
                summary_df = summarize_backtest_signals(bt_df, holding_days=holding_days)
                st.session_state.vadm_backtest = {"bt_df": bt_df, "summary_df": summary_df}

        chart_markers = []
        if st.session_state.vadm_backtest is not None:
            bt_df = st.session_state.vadm_backtest["bt_df"]
            summary_df = st.session_state.vadm_backtest["summary_df"]

            counts = bt_df["signal"].value_counts()
            bt_c1, bt_c2, bt_c3, bt_c4 = st.columns(4)
            bt_c1.metric("BUY signals", int(counts.get("BUY", 0)))
            bt_c2.metric("SELL signals", int(counts.get("SELL", 0)))
            bt_c3.metric("HOLD days", int(counts.get("HOLD", 0)))

            if not summary_df.empty:
                avg_by_signal = summary_df.groupby("signal")["forward_return_pct"].mean()
                buy_avg = avg_by_signal.get("BUY")
                bt_c4.metric(
                    f"Avg return after BUY ({holding_days}d)",
                    f"{buy_avg:+.2f}%" if buy_avg is not None else "N/A"
                )
                st.caption(
                    "This is the ACTUAL historical result, shown as-is even if it doesn't "
                    "flatter the strategy - real forward returns, not a curated summary."
                )
                st.dataframe(summary_df, use_container_width=True)
            else:
                st.info("No BUY/SELL signals fired historically with these settings.")

            chart_markers = build_signal_markers(summary_df)

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
            chart_df, height=450, title=f"{symbol}", style="Candlestick",
            overlays=overlays, markers=chart_markers,
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
# TAB: ALPHA BLACK - VADM_t, universal for whichever stock is searched above.
# ---------------------------------------------------------------------------
with tab_black:
    st.header("Alpha Black")
    st.info(
        "**VADM_t = Beta_1·V + Beta_2·D + Beta_3·(V×D)**. Everything below runs "
        "for whichever stock you searched above - not hardcoded to one stock. "
        "V uses a self-relative fallback (vs its own history) since true "
        "sector-relative V still needs live sector mapping + peer EPS data "
        "(see Sector Mapping below) - once that's set up, swap in "
        "calc_valuation_score_V() for the sector-grouped version."
    )

    st.subheader("Hypotheses")
    st.markdown("""
| | Hypothesis |
|---|---|
| **H1** | Low PE (Cheap) → higher forward returns |
| **H2** | Rising delivery pressure (Accumulation) → higher forward returns |
| **H3** | Interaction: Cheap+Accumulation together beats either alone |
| **H4** *(proposed, your call)* | Expensive+Distribution → lowest forward returns, completing the quadrant symmetrically |
""")

    if eod_df is None:
        st.warning("No price data loaded - pick a valid symbol above.")
    elif not pe_series:
        st.info("Upload the Screener Excel above (in Alpha White) to unlock V-score, "
                 "quadrant classification, and hypothesis testing here - they need annual EPS.")
    else:
        # --- Financial ratios relevant to V ---
        st.subheader("Financial Ratios (V-score inputs)")
        v_df = calc_valuation_score_V_selfrelative(
            eod_df, parsed["years"], parsed["price"], parsed["net_profit"], parsed["adjusted_shares_cr"]
        )
        latest_v_row = v_df.dropna(subset=["V"]).iloc[-1] if v_df["V"].notna().any() else None

        fr1, fr2, fr3 = st.columns(3)
        fr1.metric("Current PE (live)", f"{live_current_pe:.2f}" if live_current_pe else "N/A")
        fr1.caption("Same live PE as Alpha White (CMP ÷ latest annual EPS)")
        if latest_v_row is not None:
            fr2.metric("V-score (self-relative)", f"{latest_v_row['V']:.3f}")
            fr2.caption("0 = most expensive vs own history, 1 = cheapest")
        else:
            fr2.metric("V-score", "N/A - insufficient history")

        D_today = calc_delivery_score_D(eod_df, lookback=20).iloc[-1]
        fr3.metric("D-score (delivery, today)", f"{D_today:+.3f}" if pd.notna(D_today) else "N/A")
        fr3.caption("Simple average of 2 z-scores - not pre-fit, per your robustness decision")

        # --- Quadrant classification, today ---
        st.subheader("Quadrant Classification (today)")
        if latest_v_row is not None and pd.notna(D_today):
            today_quadrant = calc_quadrant(latest_v_row["V"], D_today)
            quadrant_colors = {
                "Cheap_Accumulation": st.success, "Expensive_Distribution": st.error,
                "Cheap_Distribution": st.warning, "Expensive_Accumulation": st.warning,
            }
            quadrant_colors.get(today_quadrant, st.info)(f"**{today_quadrant.replace('_', ' + ')}**")
        else:
            st.info("Not enough history yet to classify today's quadrant.")

        st.divider()

        # --- Hypothesis testing, button-driven, universal for this stock ---
        st.subheader("Hypothesis Testing")
        st.caption(
            "Runs H1-H4 against this stock's actual history - real result for "
            "THIS stock, not asserted as universally true across all stocks."
        )

        if "vadm_hypothesis_test" not in st.session_state:
            st.session_state.vadm_hypothesis_test = None

        forward_days_black = st.slider("Forward-return window for testing (days)", 5, 60, 20, key="black_fwd_days")

        if st.button("Test Hypotheses (H1-H4)"):
            test_result = test_vadm_t_hypotheses(
                eod_df, parsed["years"], parsed["price"], parsed["net_profit"], parsed["adjusted_shares_cr"],
                forward_days=forward_days_black,
            )
            st.session_state.vadm_hypothesis_test = test_result

        quadrant_markers = []
        if st.session_state.vadm_hypothesis_test is not None:
            test_result = st.session_state.vadm_hypothesis_test
            if "error" in test_result and not test_result.get("quadrant_stats"):
                st.warning(test_result["error"])
            else:
                st.markdown(f"**{test_result['n_total']} clean trading days tested**")

                qc1, qc2, qc3, qc4 = st.columns(4)
                for col, qname in zip([qc1, qc2, qc3, qc4],
                                       ["Cheap_Accumulation", "Cheap_Distribution",
                                        "Expensive_Accumulation", "Expensive_Distribution"]):
                    stats = test_result["quadrant_stats"].get(qname)
                    if stats:
                        col.metric(qname.replace("_", "+"), f"{stats['mean']:+.2f}%", f"n={int(stats['count'])}")
                    else:
                        col.metric(qname.replace("_", "+"), "N/A")

                st.markdown("**Verdicts (real, on this stock's history):**")
                for h, verdict in test_result["verdicts"].items():
                    (st.success if verdict == "HOLDS" else st.error)(f"{h}: **{verdict}**")

                quadrant_markers = build_quadrant_markers(test_result["detail_df"])

        st.divider()

        # --- Chart with quadrant signals ---
        st.subheader("Chart")
        st.caption("C+A markers = Cheap+Accumulation (H1-H3 predicted best). "
                    "E+D markers = Expensive+Distribution (H4 predicted worst). "
                    "Run 'Test Hypotheses' above to populate markers.")

        chart_df_black = eod_df.rename(columns={
            "Date": "date", "Open": "open", "High": "high", "Low": "low", "Close": "close", "Volume": "volume",
        }).copy()
        chart_df_black["date"] = chart_df_black["date"].dt.strftime("%Y-%m-%d")
        chart_df_black = chart_df_black.replace({np.nan: None})

        black_price_indicator = PriceIndicator(
            chart_df_black, height=400, title=f"{symbol} - VADM_t", style="Candlestick",
            markers=quadrant_markers,
        )
        black_price_indicator.calculate()

        lightweight_charts_v5_component(
            name=f"{symbol}_black", charts=[black_price_indicator.pane()], height=400,
            key=f"black_chart_{symbol}",
        )

    st.divider()
    st.subheader("Sector Mapping Setup")
    st.caption(
        "Click to fetch live from NSE - runs on Streamlit Cloud when deployed, "
        "no local code needed. Uses NSE's own sectoral indices as sector "
        "labels (your choice, revisiting the earlier idea)."
    )

    if "vadm_sector_mapping" not in st.session_state:
        st.session_state.vadm_sector_mapping = None

    if st.button("Fetch Sector Mapping (Live NSE)"):
        try:
            result = build_sector_mapping(CANDIDATE_SECTORAL_INDICES)
            st.session_state.vadm_sector_mapping = result
        except Exception as e:
            st.error(
                f"Fetch failed: {e}. If this is an auth/connection error, the "
                f"deployed app's network to nseindia.com may be blocked or rate-limited - "
                f"try again in a moment."
            )

    if st.session_state.vadm_sector_mapping is not None:
        mapping_result = st.session_state.vadm_sector_mapping
        mapping = mapping_result["mapping"]
        failures = mapping_result["failures"]

        mc1, mc2 = st.columns(2)
        mc1.metric("Symbols mapped", len(mapping))
        mc2.metric("Sectors covered", len(set(mapping.values())))

        if failures:
            st.warning(f"{len(failures)} index(es) failed to fetch: {failures}")

        if mapping:
            preview_df = pd.DataFrame(
                [{"Symbol": s, "Sector": sec} for s, sec in mapping.items()]
            ).sort_values(["Sector", "Symbol"])
            st.dataframe(preview_df, use_container_width=True, height=300)
            st.caption(
                "This mapping is now available for V-score building once you "
                "also upload multiple stocks' Screener Excel files per sector."
            )
