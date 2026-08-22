"""
NSE & BSE Deals Tracker — Investment Banking / Institutional Research Dashboard
Streamlit Main Application Entry Point
"""

import logging
from datetime import date, datetime
from typing import Dict, Any, List
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from nse_fetch import fetch_nse_deals
from bse_fetch import fetch_bse_deals
from price_fetch import get_trading_days_range, enrich_deals_with_prices
from excel_export import generate_excel_workbook

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Page configuration
st.set_page_config(
    page_title="NSE & BSE Deals Tracker",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom Institutional CSS Styling
st.markdown("""
<style>
    .main { background-color: #0f172a; }
    .stMetric {
        background-color: #1e293b;
        padding: 14px 18px;
        border-radius: 8px;
        border-left: 4px solid #0284c7;
        box-shadow: 0 1px 3px rgba(0,0,0,0.3);
    }
    .metric-card-positive { border-left-color: #10b981 !important; }
    .metric-card-negative { border-left-color: #ef4444 !important; }
    .metric-title { font-size: 0.85rem; color: #94a3b8; font-weight: 600; text-transform: uppercase; }
    .metric-value { font-size: 1.5rem; font-weight: 700; color: #f8fafc; }
    .status-badge {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 4px;
        font-size: 0.8rem;
        font-weight: 600;
        margin-right: 6px;
    }
    .badge-success { background-color: #064e3b; color: #6ee7b7; border: 1px solid #059669; }
    .badge-warning { background-color: #78350f; color: #fde68a; border: 1px solid #d97706; }
    .badge-danger { background-color: #7f1d1d; color: #fca5a5; border: 1px solid #dc2626; }
    .section-header {
        font-size: 1.15rem;
        font-weight: 700;
        color: #e2e8f0;
        border-bottom: 2px solid #334155;
        padding-bottom: 6px;
        margin-top: 24px;
        margin-bottom: 14px;
    }
    .obs-box {
        background-color: #1e293b;
        border: 1px solid #334155;
        border-radius: 8px;
        padding: 16px;
        margin-bottom: 16px;
    }
</style>
""", unsafe_allow_html=True)


@st.cache_data(ttl=3600, show_spinner=False)
def load_all_deal_data(from_d: date, to_d: date):
    """
    Parallel cached loader for NSE, BSE, and Historical Price enrichment.
    """
    nse_df, nse_status = fetch_nse_deals(from_d, to_d)
    bse_df, bse_status = fetch_bse_deals(from_d, to_d)

    frames = []
    if not nse_df.empty:
        frames.append(nse_df)
    if not bse_df.empty:
        frames.append(bse_df)

    if not frames:
        return pd.DataFrame(), nse_status, bse_status, {"status": "Unavailable", "enriched_count": 0, "unavailable_count": 0}, 0

    combined_df = pd.concat(frames, ignore_index=True)
    raw_total = len(combined_df)

    # Deduplicate exact transactions
    combined_df = combined_df.drop_duplicates(subset=["exchange", "deal_type", "date", "symbol", "client_name", "buy_sell", "quantity", "price"])
    duplicates_removed = raw_total - len(combined_df)

    # Enrich with historical prices (T, T-1, T-5, T-15)
    enriched_df, price_status = enrich_deals_with_prices(combined_df)

    return enriched_df, nse_status, bse_status, price_status, duplicates_removed


def compute_institutional_analytics(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Computes executive metrics, observations, and rankings.
    """
    stats: Dict[str, Any] = {}
    if df.empty:
        return stats

    stats["total_deals"] = len(df)
    stats["total_val_cr"] = df["deal_value_cr"].sum()

    stats["nse_val_cr"] = df[df["exchange"] == "NSE"]["deal_value_cr"].sum()
    stats["bse_val_cr"] = df[df["exchange"] == "BSE"]["deal_value_cr"].sum()

    stats["bulk_val_cr"] = df[df["deal_type"] == "Bulk"]["deal_value_cr"].sum()
    stats["block_val_cr"] = df[df["deal_type"] == "Block"]["deal_value_cr"].sum()

    stats["buy_val_cr"] = df[df["buy_sell"] == "Buy"]["deal_value_cr"].sum()
    stats["sell_val_cr"] = df[df["buy_sell"] == "Sell"]["deal_value_cr"].sum()
    stats["net_buy_cr"] = stats["buy_val_cr"] - stats["sell_val_cr"]

    # Top Stocks Table
    stock_group = df.groupby("symbol").agg(
        deal_count=("quantity", "count"),
        quantity=("quantity", "sum"),
        deal_value_cr=("deal_value_cr", "sum"),
        buy_val_cr=("deal_value_cr", lambda x: df.loc[x.index][df.loc[x.index, "buy_sell"] == "Buy"]["deal_value_cr"].sum()),
        sell_val_cr=("deal_value_cr", lambda x: df.loc[x.index][df.loc[x.index, "buy_sell"] == "Sell"]["deal_value_cr"].sum()),
        deal_date_close=("deal_date_close", "first"),
        close_vs_t1_pct=("close_vs_t1_pct", "mean"),
        close_vs_t5_pct=("close_vs_t5_pct", "mean"),
        close_vs_t15_pct=("close_vs_t15_pct", "mean"),
        deal_vs_close_pct=("deal_vs_close_pct", "mean"),
    ).reset_index()
    stock_group["net_val_cr"] = stock_group["buy_val_cr"] - stock_group["sell_val_cr"]
    stats["top_stocks_df"] = stock_group.sort_values(by="deal_value_cr", ascending=False).reset_index(drop=True)

    # Top Clients Table
    client_group = df.groupby("client_name").agg(
        deal_count=("quantity", "count"),
        unique_stocks=("symbol", "nunique"),
        deal_value_cr=("deal_value_cr", "sum"),
        buy_val_cr=("deal_value_cr", lambda x: df.loc[x.index][df.loc[x.index, "buy_sell"] == "Buy"]["deal_value_cr"].sum()),
        sell_val_cr=("deal_value_cr", lambda x: df.loc[x.index][df.loc[x.index, "buy_sell"] == "Sell"]["deal_value_cr"].sum()),
    ).reset_index()
    client_group["net_val_cr"] = client_group["buy_val_cr"] - client_group["sell_val_cr"]
    stats["top_clients_df"] = client_group.sort_values(by="deal_value_cr", ascending=False).reset_index(drop=True)

    # Factual Investment Observations
    obs: List[str] = []
    top_s = stats["top_stocks_df"].iloc[0] if not stats["top_stocks_df"].empty else None
    top_c = stats["top_clients_df"].iloc[0] if not stats["top_clients_df"].empty else None

    if top_s is not None:
        obs.append(f"{top_s['symbol']} recorded highest transaction concentration of ₹{top_s['deal_value_cr']:,.2f} Cr across {top_s['deal_count']} reported deal(s).")
    if top_c is not None and top_c['client_name'] != "NOT SPECIFIED":
        obs.append(f"{top_c['client_name']} represented the most active participant with ₹{top_c['deal_value_cr']:,.2f} Cr total transacted across {top_c['unique_stocks']} security(ies).")

    nse_share = (stats["nse_val_cr"] / stats["total_val_cr"] * 100.0) if stats["total_val_cr"] > 0 else 0.0
    bse_share = 100.0 - nse_share
    obs.append(f"NSE accounted for {nse_share:.1f}% of total reported value; BSE accounted for {bse_share:.1f}%.")

    block_share = (stats["block_val_cr"] / stats["total_val_cr"] * 100.0) if stats["total_val_cr"] > 0 else 0.0
    bulk_share = 100.0 - block_share
    obs.append(f"Block window trades comprised {block_share:.1f}% of deal volume; Bulk market deals constituted {bulk_share:.1f}%.")

    if stats["net_buy_cr"] > 0:
        obs.append(f"Aggregate net institutional inflow (reported Buy minus Sell) stood at +₹{stats['net_buy_cr']:,.2f} Cr.")
    elif stats["net_buy_cr"] < 0:
        obs.append(f"Aggregate net institutional outflow (reported Buy minus Sell) stood at -₹{abs(stats['net_buy_cr']):,.2f} Cr.")
    else:
        obs.append("Reported Buy and Sell deal volumes were evenly matched.")

    stats["observations"] = obs
    return stats


def main():
    # Sidebar
    st.sidebar.title("🏛️ Deals Tracker")
    st.sidebar.caption("Institutional Investment Research")

    st.sidebar.subheader("📅 Date Window")
    trading_days_opt = st.sidebar.selectbox("Trading Sessions", [5, 7, 10, 15, 20], index=1)
    use_custom_date = st.sidebar.checkbox("Custom Date Range", value=False)

    if use_custom_date:
        col_d1, col_d2 = st.sidebar.columns(2)
        from_date = col_d1.date_input("From", value=date.today() - pd.Timedelta(days=10))
        to_date = col_d2.date_input("To", value=date.today())
    else:
        from_date, to_date = get_trading_days_range(trading_days_opt)
        st.sidebar.info(f"Targeting last **{trading_days_opt} trading sessions** ({from_date.strftime('%d-%b-%Y')} to {to_date.strftime('%d-%b-%Y')})")

    st.sidebar.subheader("🔍 Filters")
    exch_filter = st.sidebar.selectbox("Exchange", ["All", "NSE", "BSE"], index=0)
    deal_type_filter = st.sidebar.selectbox("Deal Type", ["All", "Bulk", "Block"], index=0)
    symbol_search = st.sidebar.text_input("Filter Symbol", "").strip().upper()
    client_search = st.sidebar.text_input("Filter Client Name", "").strip().lower()

    if st.sidebar.button("🔄 Fetch Latest Data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    # Load Data
    with st.spinner("Accessing official exchange records and price histories..."):
        deals_df, nse_status, bse_status, price_status, duplicates_removed = load_all_deal_data(from_date, to_date)

    # Title & Header
    st.title("NSE & BSE Deals Tracker")
    st.caption("Investment Banking & Institutional Equities Transaction Intelligence")

    # Source Diagnostic Badges
    badge_html = "<div>"
    # NSE badge
    if nse_status["status"] == "Loaded":
        badge_html += '<span class="status-badge badge-success">NSE: Loaded</span>'
    elif nse_status["status"] == "Warning":
        badge_html += f'<span class="status-badge badge-warning">NSE: {nse_status["message"]}</span>'
    else:
        badge_html += '<span class="status-badge badge-danger">NSE: Failed</span>'

    # BSE badge
    if bse_status["status"] == "Loaded":
        badge_html += '<span class="status-badge badge-success">BSE: Loaded</span>'
    elif bse_status["status"] == "Warning":
        badge_html += f'<span class="status-badge badge-warning">BSE: {bse_status["message"]}</span>'
    else:
        badge_html += '<span class="status-badge badge-danger">BSE: Failed</span>'

    # Price status
    if price_status["status"] == "Loaded":
        badge_html += '<span class="status-badge badge-success">Price Series: Enriched</span>'
    elif price_status["status"] == "Partial":
        badge_html += '<span class="status-badge badge-warning">Price Series: Partial</span>'
    else:
        badge_html += '<span class="status-badge badge-danger">Price Series: Unavailable</span>'

    badge_html += "</div>"
    st.markdown(badge_html, unsafe_allow_html=True)
    st.write("")

    # Empty State Handling
    if deals_df.empty:
        st.warning(
            "Couldn't fetch data — NSE/BSE may be rate-limiting, the market may be closed, the selected period may "
            "contain no reported deals, or an exchange endpoint may have changed.\n\n"
            "Try **Fetch Latest Data** again or check the official exchange reports."
        )
        return

    # Apply Sidebar Table Filters
    filtered_df = deals_df.copy()
    if exch_filter != "All":
        filtered_df = filtered_df[filtered_df["exchange"] == exch_filter]
    if deal_type_filter != "All":
        filtered_df = filtered_df[filtered_df["deal_type"] == deal_type_filter]
    if symbol_search:
        filtered_df = filtered_df[filtered_df["symbol"].str.contains(symbol_search, na=False)]
    if client_search:
        filtered_df = filtered_df[filtered_df["client_name"].str.lower().str.contains(client_search, na=False)]

    # Analytics
    stats = compute_institutional_analytics(filtered_df)

    # 1. EXECUTIVE METRICS CARDS
    st.markdown('<div class="section-header">📊 Executive Deal Flow Summary</div>', unsafe_allow_html=True)
    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("Total Deals", f"{stats.get('total_deals', 0):,}")
    m2.metric("Total Value", f"₹{stats.get('total_val_cr', 0.0):,.1f} Cr")
    m3.metric("NSE Value", f"₹{stats.get('nse_val_cr', 0.0):,.1f} Cr")
    m4.metric("BSE Value", f"₹{stats.get('bse_val_cr', 0.0):,.1f} Cr")
    m5.metric("Block Deals", f"₹{stats.get('block_val_cr', 0.0):,.1f} Cr")
    net_val = stats.get('net_buy_cr', 0.0)
    m6.metric("Net Buy / (Sell)", f"₹{net_val:,.1f} Cr", delta=f"{net_val:,.1f} Cr")

    # 2. KEY DEAL FLOW OBSERVATIONS
    st.markdown('<div class="section-header">💡 Key Deal Flow Observations</div>', unsafe_allow_html=True)
    obs_list = stats.get("observations", [])
    obs_md = "<div class='obs-box'>" + "".join([f"<p style='margin-bottom: 6px;'>• <strong>{o}</strong></p>" for o in obs_list]) + "</div>"
    st.markdown(obs_md, unsafe_allow_html=True)

    # Excel Download Button
    excel_data = generate_excel_workbook(filtered_df, stats)
    st.download_button(
        label="📥 Download Institutional Excel Report (NSE_BSE_Deals_Tracker.xlsx)",
        data=excel_data,
        file_name="NSE_BSE_Deals_Tracker.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    # 3. INTERACTIVE CHARTS
    st.markdown('<div class="section-header">📈 Market Visualizations</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)

    top_stocks_plot = stats.get("top_stocks_df", pd.DataFrame()).head(10)
    if not top_stocks_plot.empty:
        fig_stock = px.bar(
            top_stocks_plot,
            x="deal_value_cr",
            y="symbol",
            orientation="h",
            title="Top 10 Stocks by Reported Deal Value (₹ Cr)",
            labels={"deal_value_cr": "Deal Value (₹ Cr)", "symbol": "Stock"},
            color="deal_value_cr",
            color_continuous_scale="Blues",
        )
        fig_stock.update_layout(yaxis={"categoryorder": "total ascending"}, template="plotly_dark", height=360)
        c1.plotly_chart(fig_stock, use_container_width=True)

    top_clients_plot = stats.get("top_clients_df", pd.DataFrame()).head(10)
    if not top_clients_plot.empty:
        fig_client = px.bar(
            top_clients_plot,
            x="deal_value_cr",
            y="client_name",
            orientation="h",
            title="Top 10 Clients by Reported Deal Value (₹ Cr)",
            labels={"deal_value_cr": "Deal Value (₹ Cr)", "client_name": "Client"},
            color="deal_value_cr",
            color_continuous_scale="Teal",
        )
        fig_client.update_layout(yaxis={"categoryorder": "total ascending"}, template="plotly_dark", height=360)
        c2.plotly_chart(fig_client, use_container_width=True)

    c3, c4, c5 = st.columns(3)
    # Exchange Split
    fig_exch = px.pie(
        filtered_df,
        names="exchange",
        values="deal_value_cr",
        title="Exchange Distribution",
        hole=0.45,
        color_discrete_sequence=["#0284c7", "#f59e0b"],
    )
    fig_exch.update_layout(template="plotly_dark", height=280)
    c3.plotly_chart(fig_exch, use_container_width=True)

    # Deal Type Split
    fig_dtype = px.pie(
        filtered_df,
        names="deal_type",
        values="deal_value_cr",
        title="Deal Type Distribution",
        hole=0.45,
        color_discrete_sequence=["#10b981", "#6366f1"],
    )
    fig_dtype.update_layout(template="plotly_dark", height=280)
    c4.plotly_chart(fig_dtype, use_container_width=True)

    # Buy vs Sell
    fig_bs = px.pie(
        filtered_df,
        names="buy_sell",
        values="deal_value_cr",
        title="Buy vs Sell Volume",
        hole=0.45,
        color_discrete_sequence=["#22c55e", "#ef4444"],
    )
    fig_bs.update_layout(template="plotly_dark", height=280)
    c5.plotly_chart(fig_bs, use_container_width=True)

    # 4. MAIN ENRICHED TRANSACTION TABLE
    st.markdown('<div class="section-header">📋 Complete Deal Ledger with Historical Price Context</div>', unsafe_allow_html=True)
    
    display_df = filtered_df[[
        "date", "exchange", "deal_type", "symbol", "security_name", "client_name",
        "buy_sell", "quantity", "price", "deal_value_cr", "deal_date_close",
        "t1_close", "t5_close", "t15_close", "deal_vs_close_pct",
        "close_vs_t1_pct", "close_vs_t5_pct", "close_vs_t15_pct"
    ]].copy()

    st.dataframe(
        display_df.style.format({
            "quantity": "{:,.0f}",
            "price": "₹{:,.2f}",
            "deal_value_cr": "₹{:,.2f} Cr",
            "deal_date_close": "₹{:,.2f}",
            "t1_close": "₹{:,.2f}",
            "t5_close": "₹{:,.2f}",
            "t15_close": "₹{:,.2f}",
            "deal_vs_close_pct": "{:+.2f}%",
            "close_vs_t1_pct": "{:+.2f}%",
            "close_vs_t5_pct": "{:+.2f}%",
            "close_vs_t15_pct": "{:+.2f}%",
        }, na_rep="-"),
        use_container_width=True,
        height=450
    )

    # 5. AGGREGATE BREAKDOWNS
    tab1, tab2, tab3, tab4 = st.tabs(["Top Stocks", "Top Clients", "Repeat Activity", "Data Quality"])

    with tab1:
        st.subheader("Top Stocks by Value")
        top_stocks_view = stats.get("top_stocks_df", pd.DataFrame())
        if not top_stocks_view.empty:
            st.dataframe(
                top_stocks_view.style.format({
                    "deal_count": "{:,}",
                    "quantity": "{:,.0f}",
                    "buy_val_cr": "₹{:,.2f} Cr",
                    "sell_val_cr": "₹{:,.2f} Cr",
                    "deal_value_cr": "₹{:,.2f} Cr",
                    "net_val_cr": "₹{:,.2f} Cr",
                    "deal_date_close": "₹{:,.2f}",
                    "close_vs_t1_pct": "{:+.2f}%",
                    "close_vs_t5_pct": "{:+.2f}%",
                    "close_vs_t15_pct": "{:+.2f}%",
                    "deal_vs_close_pct": "{:+.2f}%",
                }, na_rep="-"),
                use_container_width=True
            )

    with tab2:
        st.subheader("Top Clients by Value")
        top_clients_view = stats.get("top_clients_df", pd.DataFrame())
        if not top_clients_view.empty:
            st.dataframe(
                top_clients_view.style.format({
                    "deal_count": "{:,}",
                    "unique_stocks": "{:,}",
                    "buy_val_cr": "₹{:,.2f} Cr",
                    "sell_val_cr": "₹{:,.2f} Cr",
                    "deal_value_cr": "₹{:,.2f} Cr",
                    "net_val_cr": "₹{:,.2f} Cr",
                }, na_rep="-"),
                use_container_width=True
            )

    with tab3:
        st.subheader("Repeat Active Clients")
        repeat_clients = stats.get("top_clients_df", pd.DataFrame())
        if not repeat_clients.empty:
            repeat_clients_filtered = repeat_clients[repeat_clients["deal_count"] > 1]
            st.dataframe(
                repeat_clients_filtered.style.format({
                    "deal_count": "{:,}",
                    "unique_stocks": "{:,}",
                    "buy_val_cr": "₹{:,.2f} Cr",
                    "sell_val_cr": "₹{:,.2f} Cr",
                    "deal_value_cr": "₹{:,.2f} Cr",
                    "net_val_cr": "₹{:,.2f} Cr",
                }, na_rep="-"),
                use_container_width=True
            )

    with tab4:
        st.subheader("Data Quality & Pipeline Diagnostics")
        dq1, dq2, dq3, dq4 = st.columns(4)
        dq1.metric("Raw Deals Fetched", f"{len(deals_df) + duplicates_removed:,}")
        dq2.metric("Clean Records", f"{len(deals_df):,}")
        dq3.metric("Duplicates Deduplicated", f"{duplicates_removed:,}")
        dq4.metric("Prices Enriched", f"{price_status['enriched_count']:,}")

        st.json({
            "NSE Status": nse_status,
            "BSE Status": bse_status,
            "Price Engine Status": price_status,
            "Unique Securities Tracked": deals_df["symbol"].nunique(),
            "Min Deal Date": str(deals_df["date"].min()),
            "Max Deal Date": str(deals_df["date"].max()),
        })

    # Institutional Disclaimers
    with st.expander("ℹ️ Methodology, Data Sources & Regulatory Disclaimers"):
        st.markdown("""
        - **Data Ingestion**: Deal records are ingested from official NSE Large Deals and BSE Bulk/Block disclosures.
        - **Historical Price Enrichment**: Closing prices are derived from daily official exchange closing quotes. Price changes (1D, 5D, 15D) are computed against actual preceding trading sessions (skipping weekends and market holidays).
        - **Terminology Notice**: Premium/Discount is calculated strictly against Deal Date Market Close: `(Deal Price / Deal Date Close - 1) * 100`. Deal Price is the reported transaction rate and differs from the market closing quote.
        - **Regulatory Disclaimer**: Bulk/Block transactions represent exchange disclosures of large trades and do not alone constitute investment intent or forward recommendations. Exchange interfaces may include unofficial/reverse-engineered endpoints and may change without notice. Verify data against official circulars before material capital deployment.
        """)


if __name__ == "__main__":
    main()
