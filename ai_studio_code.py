import streamlit as st
import pandas as pd
import requests
import datetime
from datetime import timedelta
import io

# --- PAGE CONFIG ---
st.set_page_config(page_title="NSE/BSE Deals Tracker", layout="wide", page_icon="📈")

# --- CUSTOM CSS FOR INSTITUTIONAL LOOK ---
st.markdown("""
    <style>
    .main { background-color: #F8F9FA; }
    .stMetric { background-color: #ffffff; border: 1px solid #DDE1E6; padding: 15px; border-radius: 5px; }
    .stButton>button { width: 100%; border-radius: 5px; background-color: #1B2631; color: white; }
    </style>
    """, unsafe_allow_html=True)

# --- API FETCHING LOGIC ---
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.nseindia.com/"
}

def fetch_nse_data(deal_type="bulk", days=10):
    end_date = datetime.date.today()
    start_date = end_date - timedelta(days=days)
    
    url = f"https://www.nseindia.com/api/historical/{deal_type}-deals"
    params = {"from": start_date.strftime("%d-%m-%Y"), "to": end_date.strftime("%d-%m-%Y")}
    
    session = requests.Session()
    session.get("https://www.nseindia.com", headers=HEADERS, timeout=10) # Get cookies
    response = session.get(url, params=params, headers=HEADERS, timeout=10)
    
    if response.status_code == 200:
        data = response.json().get('data', [])
        df = pd.DataFrame(data)
        if not df.empty:
            df = df.rename(columns={
                'CH_SYMBOL': 'Symbol', 'CH_SYMBOL_NAME': 'Security Name', 
                'CH_CLIENT_NAME': 'Client Name', 'CH_TRADE_TYPE': 'Buy/Sell',
                'CH_QUANTITY': 'Quantity', 'CH_PRICE': 'Price', 'CH_TIMESTAMP': 'Date'
            })
            df['Exchange'] = 'NSE'
            df['Deal Type'] = deal_type.capitalize()
            return df[['Exchange', 'Deal Type', 'Date', 'Symbol', 'Security Name', 'Client Name', 'Buy/Sell', 'Quantity', 'Price']]
    return pd.DataFrame()

def fetch_bse_data(deal_type="B", days=10):
    end_date = datetime.date.today()
    start_date = end_date - timedelta(days=days)
    url = "https://api.bseindia.com/BseIndiaAPI/api/BulkBlockDeal/w"
    params = {
        "flag": "1", "fdate": start_date.strftime("%Y%m%d"), 
        "todate": end_date.strftime("%Y%m%d"), "deal_type": deal_type
    }
    try:
        response = requests.get(url, params=params, headers=HEADERS, timeout=10)
        data = response.json()
        df = pd.DataFrame(data['Table'] if isinstance(data, dict) else data)
        if not df.empty:
            df = df.rename(columns={
                'Symbol': 'Symbol', 'ScripName': 'Security Name', 'ClientName': 'Client Name',
                'BuySell': 'Buy/Sell', 'Qty': 'Quantity', 'Rate': 'Price', 'DealDate': 'Date'
            })
            df['Exchange'] = 'BSE'
            df['Deal Type'] = 'Bulk' if deal_type == "B" else 'Block'
            df['Buy/Sell'] = df['Buy/Sell'].replace({'B': 'Buy', 'S': 'Sell'})
            return df[['Exchange', 'Deal Type', 'Date', 'Symbol', 'Security Name', 'Client Name', 'Buy/Sell', 'Quantity', 'Price']]
    except: pass
    return pd.DataFrame()

# --- EXCEL GENERATION ENGINE ---
def generate_excel(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book
    
    # Pre-calculate data for Analysis
    df['Date'] = pd.to_datetime(df['Date']).dt.date
    df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce')
    df['Price'] = pd.to_numeric(df['Price'], errors='coerce')
    df['Deal Value'] = df['Quantity'] * df['Price']
    df['Deal Value (₹ Cr)'] = df['Deal Value'] / 10_000_000
    df = df.dropna(subset=['Deal Value'])

    # Formats
    header_fmt = workbook.add_format({'bg_color': '#1B2631', 'font_color': 'white', 'bold': True, 'border': 1})
    cr_fmt = workbook.add_format({'num_format': '₹#,##0.00 "Cr"', 'align': 'right'})
    num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'right'})
    date_fmt = workbook.add_format({'num_format': 'dd-mmm-yyyy', 'align': 'center'})

    # 1. RAW DATA SHEET
    df.to_excel(writer, sheet_name='Raw Data', index=False)
    worksheet = writer.sheets['Raw Data']
    worksheet.add_table(0, 0, len(df), len(df.columns)-1, {'name': 'Deals_Data', 'style': 'TableStyleMedium2'})
    worksheet.set_column('K:K', 15, cr_fmt)
    worksheet.set_column('C:C', 12, date_fmt)

    # 2. PIVOTS SHEET (Structured Data for user)
    p_sheet = workbook.add_worksheet('Pivots')
    sym_pivot = df.groupby('Symbol')['Deal Value (₹ Cr)'].sum().sort_values(ascending=False).head(20)
    sym_pivot.to_excel(writer, sheet_name='Pivots', startrow=2)

    # 3. ANALYSIS SHEET (Dashboard)
    dash = workbook.add_worksheet('Analysis')
    dash.set_column('B:L', 18)
    dash.write('B2', 'EXECUTIVE DEALS TRACKER', workbook.add_format({'bold': True, 'font_size': 18}))
    
    # KPI Formulas
    dash.write('B5', 'Total Deals', header_fmt)
    dash.write_formula('B6', '=COUNT(Deals_Data[Symbol])', num_fmt)
    dash.write('C5', 'Total Value (₹ Cr)', header_fmt)
    dash.write_formula('C6', '=SUM(Deals_Data[Deal Value (₹ Cr)])', cr_fmt)
    
    # Top 10 Deals Table
    dash.write('B10', 'TOP 10 INDIVIDUAL DEALS', header_fmt)
    top_10 = df.nlargest(10, 'Deal Value (₹ Cr)')
    top_10[['Date', 'Symbol', 'Client Name', 'Buy/Sell', 'Deal Value (₹ Cr)']].to_excel(writer, sheet_name='Analysis', startrow=11, startcol=1, index=False)

    # Add Chart
    chart = workbook.add_chart({'type': 'bar'})
    chart.add_series({
        'categories': ['Pivots', 3, 0, 13, 0],
        'values':     ['Pivots', 3, 1, 13, 1],
        'name': 'Value (₹ Cr)'
    })
    dash.insert_chart('H10', chart)

    writer.close()
    return output.getvalue()

# --- STREAMLIT UI ---
st.title("🏛️ NSE & BSE Deals Tracker")
st.subheader("Investment Banking / Institutional Research Dashboard")

with st.sidebar:
    st.header("Settings")
    lookback = st.slider("Lookback Days", 1, 15, 7)
    fetch_btn = st.button("Fetch Latest Data")
    st.info("Fetches Bulk & Block deals from NSE and BSE official APIs.")

if fetch_btn:
    with st.spinner("Accessing Exchange APIs..."):
        # Fetching
        n_bulk = fetch_nse_data("bulk", lookback)
        n_block = fetch_nse_data("block", lookback)
        b_bulk = fetch_bse_data("B", lookback)
        b_block = fetch_bse_data("K", lookback)
        
        combined_df = pd.concat([n_bulk, n_block, b_bulk, b_block], ignore_index=True)
        
        if not combined_df.empty:
            combined_df['Quantity'] = pd.to_numeric(combined_df['Quantity'], errors='coerce')
            combined_df['Price'] = pd.to_numeric(combined_df['Price'], errors='coerce')
            combined_df['Value_Cr'] = (combined_df['Quantity'] * combined_df['Price']) / 10_000_000
            
            # Dashboard Metrics
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total Deals", len(combined_df))
            col2.metric("Total Value", f"₹{combined_df['Value_Cr'].sum():,.2f} Cr")
            col3.metric("NSE Value", f"₹{combined_df[combined_df['Exchange']=='NSE']['Value_Cr'].sum():,.2f} Cr")
            col4.metric("BSE Value", f"₹{combined_df[combined_df['Exchange']=='BSE']['Value_Cr'].sum():,.2f} Cr")

            # Chart
            st.write("### Top 15 Stocks by Deal Value")
            top_stocks = combined_df.groupby('Symbol')['Value_Cr'].sum().sort_values(ascending=True).tail(15)
            st.bar_chart(top_stocks)

            # Table Preview
            st.write("### Recent Deal Preview")
            st.dataframe(combined_df.sort_values(by='Date', ascending=False).head(20), use_container_width=True)

            # Download Button
            excel_data = generate_excel(combined_df)
            st.download_button(
                label="📥 Download Production-Ready Excel Workbook",
                data=excel_data,
                file_name=f"NSE_BSE_Deals_{datetime.date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No data found. The exchange APIs might be rate-limiting or the markets are closed.")

st.markdown("---")
st.caption("Disclaimer: This tool uses unofficial API endpoints. Material decisions should be verified with official exchange PDFs.")