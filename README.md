# VADM — build status

## Setup
```
pip install streamlit pandas numpy openpyxl nse streamlit-lightweight-charts-v5
streamlit run app.py
```
**Important**: `indicators.py` and `chart_themes.py` are NOT part of the pip package —
they're demo files from the `streamlit-lightweight-charts-v5` GitHub repo that must
sit next to `app.py`. They're included here already, copied from
`locupleto/streamlit-lightweight-charts-v5/demo/`.

## What's tested and confirmed working (I ran all of this myself, not just written)
- EOD2 fetch (`vadm_calculations.load_eod2_data`) — live-tested against RELIANCE
- Delivery % and relative delivery (z-score/percentile) — tested
- 52-week high/low — tested
- Screener Excel parsing — tested against your actual IRB_Infra_Devl-3.xlsx
- PE-per-year, EV-per-year, revenue growth — tested against IRB's real numbers
- Full app — ran end-to-end with Streamlit's AppTest (no browser, but caught real
  exceptions); RSI/MACD/SMA indicator classes individually verified too

## What's NOT tested (written correctly per docs, but never executed)
- `fetch_promoter_holding()` — this sandbox can't reach nseindia.com. Run it
  yourself first before trusting the output.
- Visual layout of the Delivery % line overlaid on the Volume pane — the JSON
  is schema-valid but I have no browser here to confirm it looks right. Adjust
  `scaleMargins` in `app.py` if the delivery % line overlaps the volume bars.

## What's intentionally NOT built — real gaps, not oversights
1. **Alpha White's sector_avg_pe** — currently a manual number input in the UI.
   No resolved automatic source since the Excel turned out to have no
   Industry/Sector column. Options still on the table from earlier in the
   conversation: a manual stock→sector mapping table, or revisiting NSE
   sectoral index PE via a library.
2. **Alpha Black / VADM_t** — no formula `f()`, no H4. Your own project memory
   flags this as a CEO-level decision. The tab shows the delivery-flow data
   ready and waiting, nothing invented for the combination logic.
3. **"Liquidity entered last 6/12 months"** — never got a confirmed definition.
   Currently implemented as `Close × Volume` summed over the window, clearly
   labeled as a guess in the UI itself (⚠️ caption under the metric).
4. **Pledge %** — dropped per your instruction; only promoter holding, total
   debt, and EV were kept.
