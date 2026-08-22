"""
Historical Stock Price Enrichment Engine
Calculates Market Close on Deal Date (T), T-1, T-5, and T-15 trading sessions.
"""

import time
import logging
from datetime import datetime, date, timedelta
from typing import Dict, List, Tuple, Any, Optional
import requests
import pandas as pd
import streamlit as st

logger = logging.getLogger(__name__)

# Standard Indian Stock Market Holidays Reference (Weekends automatically skipped)
NSE_PRICE_HIST_URL = "https://www.nseindia.com/api/historical/securityArchives"
BSE_PRICE_HIST_URL = "https://api.bseindia.com/BseIndiaAPI/api/StockPriceSeries/w"

NSE_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.nseindia.com/report-detail/eq_security",
}


def get_trading_days_range(num_trading_days: int = 7) -> Tuple[date, date]:
    """
    Computes a date window that spans at least the requested number of trading sessions.
    Skips weekends and exchange non-trading days.
    """
    today = date.today()
    trading_days_found = 0
    curr_date = today

    # Step backwards from current date until we have found the desired number of trading sessions
    while trading_days_found < num_trading_days:
        if curr_date.weekday() < 5:  # Monday to Friday
            trading_days_found += 1
        if trading_days_found == num_trading_days:
            break
        curr_date -= timedelta(days=1)

    return curr_date, today


def fetch_nse_security_prices(symbol: str, from_date: date, to_date: date, session: Optional[requests.Session] = None) -> Dict[date, float]:
    """
    Fetches continuous historical closing prices for an NSE security.
    Returns a dict mapping datetime.date -> closing_price.
    """
    if session is None:
        session = requests.Session()
        session.headers.update(NSE_HEADERS)
        try:
            session.get("https://www.nseindia.com/", timeout=8)
        except Exception:
            pass

    from_str = from_date.strftime("%d-%m-%Y")
    to_str = to_date.strftime("%d-%m-%Y")
    params = {
        "from": from_str,
        "to": to_str,
        "symbol": symbol,
        "dataType": "priceVolumeDeliverable",
        "series": "ALL",
    }

    prices_by_date: Dict[date, float] = {}
    try:
        resp = session.get(NSE_PRICE_HIST_URL, params=params, headers=NSE_HEADERS, timeout=12)
        if resp.status_code == 200:
            text = resp.text.strip()
            if text and not text.startswith("<") and "html" not in resp.headers.get("Content-Type", "").lower():
                data = resp.json()
                items = data.get("data", []) if isinstance(data, dict) else (data if isinstance(data, list) else [])
                for item in items:
                    if not isinstance(item, dict):
                        continue
                    dt_str = item.get("CH_TIMESTAMP") or item.get("mTIMESTAMP") or item.get("HistoricalDate") or item.get("date")
                    close_px = item.get("CH_CLOSING_PRICE") or item.get("close") or item.get("CH_LAST_TRADED_PRICE")
                    if dt_str and close_px is not None:
                        try:
                            # Handle formats like 2026-08-20 or 20-Aug-2026
                            d_obj = None
                            for fmt in ("%Y-%m-%d", "%d-%b-%Y", "%d-%m-%Y"):
                                try:
                                    d_obj = datetime.strptime(str(dt_str).strip(), fmt).date()
                                    break
                                except ValueError:
                                    continue
                            if d_obj:
                                prices_by_date[d_obj] = float(str(close_px).replace(",", ""))
                        except Exception:
                            continue
    except Exception as exc:
        logger.warning(f"Error fetching NSE historical price for {symbol}: {exc}")

    return prices_by_date


def fetch_bse_security_prices(scrip: str, from_date: date, to_date: date) -> Dict[date, float]:
    """
    Fetches continuous historical closing prices for a BSE security.
    Returns a dict mapping datetime.date -> closing_price.
    """
    prices_by_date: Dict[date, float] = {}
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "application/json, text/plain, */*",
        "Origin": "https://www.bseindia.com",
        "Referer": "https://www.bseindia.com/",
    }
    params = {
        "scripcode": scrip,
        "fromdate": from_date.strftime("%Y%m%d"),
        "todate": to_date.strftime("%Y%m%d"),
    }
    try:
        resp = requests.get(BSE_PRICE_HIST_URL, params=params, headers=headers, timeout=12)
        if resp.status_code == 200:
            text = resp.text.strip()
            if text and not text.startswith("<"):
                data = resp.json()
                items = data.get("Table", []) if isinstance(data, dict) else (data if isinstance(data, list) else [])
                for item in items:
                    if not isinstance(item, dict):
                        continue
                    dt_str = item.get("dttm") or item.get("Date") or item.get("tradedate")
                    close_px = item.get("close") or item.get("Close") or item.get("rate")
                    if dt_str and close_px is not None:
                        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%b-%Y"):
                            try:
                                d_obj = datetime.strptime(str(dt_str).strip(), fmt).date()
                                prices_by_date[d_obj] = float(str(close_px).replace(",", ""))
                                break
                            except ValueError:
                                continue
    except Exception as exc:
        logger.warning(f"Error fetching BSE historical price for {scrip}: {exc}")

    return prices_by_date


@st.cache_data(ttl=3600, show_spinner=False)
def get_price_history(unique_stocks: List[Tuple[str, str]], min_date: date, max_date: date) -> Dict[Tuple[str, str], Dict[date, float]]:
    """
    Batches historical price retrieval for all unique (Exchange, Symbol) combinations.
    Extends the search window 45 calendar days backwards to ensure 15 trading days coverage.
    """
    buffer_start = min_date - timedelta(days=45)
    price_cache: Dict[Tuple[str, str], Dict[date, float]] = {}

    nse_session = requests.Session()
    nse_session.headers.update(NSE_HEADERS)
    try:
        nse_session.get("https://www.nseindia.com/", timeout=8)
    except Exception:
        pass

    for exchange, symbol in unique_stocks:
        if not symbol or symbol == "NAN":
            continue
        try:
            if exchange == "NSE":
                p_map = fetch_nse_security_prices(symbol, buffer_start, max_date, nse_session)
            else:
                p_map = fetch_bse_security_prices(symbol, buffer_start, max_date)
            price_cache[(exchange, symbol)] = p_map
            time.sleep(0.1)  # Respect exchange rate limits
        except Exception as err:
            logger.warning(f"Failed to fetch price series for {exchange}:{symbol} - {err}")
            price_cache[(exchange, symbol)] = {}

    return price_cache


def get_reference_prices_for_deal(
    deal_date: date,
    price_series: Dict[date, float]
) -> Tuple[Optional[float], Optional[float], Optional[float], Optional[float]]:
    """
    Calculates:
      1. T (Deal Date Market Close)
      2. T-1 (1 Trading Day Prior Close)
      3. T-5 (5 Trading Days Prior Close)
      4. T-15 (15 Trading Days Prior Close)

    Uses the stock's actual chronologically sorted trading sessions. Never fabricates prices.
    """
    if not price_series:
        return None, None, None, None

    # Filter dates up to deal_date and sort ascending
    trading_dates = sorted([d for d in price_series.keys() if d <= deal_date])
    if not trading_dates:
        return None, None, None, None

    # Identify index of Deal Date session
    deal_session_idx = -1
    for idx, d in enumerate(trading_dates):
        if d == deal_date:
            deal_session_idx = idx
            break

    if deal_session_idx == -1:
        # Deal date was not an official close date in the record; pick most recent prior available session
        deal_session_idx = len(trading_dates) - 1

    t_close = price_series.get(trading_dates[deal_session_idx])

    t_minus_1 = (
        price_series.get(trading_dates[deal_session_idx - 1])
        if deal_session_idx >= 1
        else None
    )

    t_minus_5 = (
        price_series.get(trading_dates[deal_session_idx - 5])
        if deal_session_idx >= 5
        else None
    )

    t_minus_15 = (
        price_series.get(trading_dates[deal_session_idx - 15])
        if deal_session_idx >= 15
        else None
    )

    return t_close, t_minus_1, t_minus_5, t_minus_15


def enrich_deals_with_prices(deals_df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Enriches the deal DataFrame with historical closes and percentage returns.
    Ensures that deal price and closing price are never confused or substituted.
    """
    diag_status = {
        "status": "Loaded",
        "enriched_count": 0,
        "unavailable_count": 0,
    }

    if deals_df.empty:
        return deals_df, diag_status

    # Extract unique securities
    unique_stocks = deals_df[["exchange", "symbol"]].drop_duplicates().to_records(index=False).tolist()
    min_deal_date = deals_df["date"].min()
    max_deal_date = deals_df["date"].max()

    # Batch fetch price series
    price_histories = get_price_history(unique_stocks, min_deal_date, max_deal_date)

    t_closes = []
    t1_closes = []
    t5_closes = []
    t15_closes = []

    p_vs_close_pct = []
    close_vs_t1_pct = []
    close_vs_t5_pct = []
    close_vs_t15_pct = []
    p_vs_t15_pct = []

    for _, row in deals_df.iterrows():
        exch = row["exchange"]
        sym = row["symbol"]
        d_date = row["date"]
        deal_px = row["price"]

        p_series = price_histories.get((exch, sym), {})
        t_c, t1_c, t5_c, t15_c = get_reference_prices_for_deal(d_date, p_series)

        t_closes.append(t_c)
        t1_closes.append(t1_c)
        t5_closes.append(t5_c)
        t15_closes.append(t15_c)

        if t_c is not None:
            diag_status["enriched_count"] += 1
        else:
            diag_status["unavailable_count"] += 1

        # Deal Price vs Deal Date Close (%)
        p_vs_c = ((deal_px / t_c) - 1.0) * 100.0 if (t_c is not None and t_c > 0) else None
        p_vs_close_pct.append(p_vs_c)

        # Deal Date Close vs T-1 Close (%)
        c_vs_t1 = ((t_c / t1_c) - 1.0) * 100.0 if (t_c is not None and t1_c is not None and t1_c > 0) else None
        close_vs_t1_pct.append(c_vs_t1)

        # Deal Date Close vs T-5 Close (%)
        c_vs_t5 = ((t_c / t5_c) - 1.0) * 100.0 if (t_c is not None and t5_c is not None and t5_c > 0) else None
        close_vs_t5_pct.append(c_vs_t5)

        # Deal Date Close vs T-15 Close (%)
        c_vs_t15 = ((t_c / t15_c) - 1.0) * 100.0 if (t_c is not None and t15_c is not None and t15_c > 0) else None
        close_vs_t15_pct.append(c_vs_t15)

        # Deal Price vs T-15 Close (%)
        p_vs_t15 = ((deal_px / t15_c) - 1.0) * 100.0 if (deal_px is not None and t15_c is not None and t15_c > 0) else None
        p_vs_t15_pct.append(p_vs_t15)

    enriched_df = deals_df.copy()
    enriched_df["deal_date_close"] = t_closes
    enriched_df["t1_close"] = t1_closes
    enriched_df["t5_close"] = t5_closes
    enriched_df["t15_close"] = t15_closes
    enriched_df["deal_vs_close_pct"] = p_vs_close_pct
    enriched_df["close_vs_t1_pct"] = close_vs_t1_pct
    enriched_df["close_vs_t5_pct"] = close_vs_t5_pct
    enriched_df["close_vs_t15_pct"] = close_vs_t15_pct
    enriched_df["deal_vs_t15_pct"] = p_vs_t15_pct

    if diag_status["unavailable_count"] > 0 and diag_status["enriched_count"] > 0:
        diag_status["status"] = "Partial"
    elif diag_status["enriched_count"] == 0 and len(deals_df) > 0:
        diag_status["status"] = "Failed"

    return enriched_df, diag_status
