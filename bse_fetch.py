"""
BSE Bulk and Block Deals Fetcher with JSON/Table Response Normalization
"""

import time
import logging
from datetime import datetime, date
from typing import Dict, List, Tuple, Any, Optional
import requests
import pandas as pd

logger = logging.getLogger(__name__)

# ==============================================================================
# BSE UNOFFICIAL ENDPOINT DISCLAIMER:
# These are unofficial/reverse-engineered endpoints, similar in approach to open-source
# BseIndiaApi implementations, and may change without notice. The application fails
# gracefully and returns an empty dataset with diagnostic warnings if BSE blocks or changes schema.
# ==============================================================================

BSE_API_URL = "https://api.bseindia.com/BseIndiaAPI/api/BulkBlockDeal/w"
BSE_REFERER_BULK = "https://www.bseindia.com/markets/equity/EQReports/bulk_deals.aspx"
BSE_ORIGIN = "https://www.bseindia.com"

BSE_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Origin": BSE_ORIGIN,
    "Referer": BSE_REFERER_BULK,
    "Connection": "keep-alive",
}


def _parse_bse_number(val: Any) -> float:
    """Safely converts numeric or formatted string fields to float."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    try:
        s = str(val).replace(",", "").replace("₹", "").strip()
        return float(s) if s else 0.0
    except Exception:
        return 0.0


def _parse_bse_date(val: Any) -> Optional[date]:
    """Converts diverse BSE date strings to date objects."""
    if not val:
        return None
    val_str = str(val).strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d-%b-%Y", "%Y%m%d"):
        try:
            return datetime.strptime(val_str, fmt).date()
        except ValueError:
            continue
    return None


def fetch_bse_endpoint(deal_type_param: str, from_date_str: str, to_date_str: str) -> Tuple[List[Dict[str, Any]], str]:
    """
    Safely queries the BSE BulkBlockDeal endpoint and handles varied JSON payloads.
    Never blindly does response.json().
    """
    params = {
        "flag": "1",
        "fdate": from_date_str,
        "todate": to_date_str,
        "scripcode": "",
        "smcode": "",
        "deal_type": deal_type_param,  # 'B' for Bulk, 'K' for Block
    }

    retries = 3
    backoff = [1.5, 3.0, 5.0]

    for attempt in range(retries):
        try:
            resp = requests.get(BSE_API_URL, headers=BSE_HEADERS, params=params, timeout=15)
            if resp.status_code == 403:
                time.sleep(backoff[attempt])
                continue
            if resp.status_code != 200:
                time.sleep(backoff[attempt])
                continue

            text = resp.text.strip()
            if not text:
                return [], "BSE returned empty response."

            content_type = resp.headers.get("Content-Type", "").lower()
            if text.startswith("<") or "html" in content_type:
                time.sleep(backoff[attempt])
                continue

            try:
                data = resp.json()
            except Exception as json_err:
                logger.warning(f"BSE JSON parse error: {json_err}")
                time.sleep(backoff[attempt])
                continue

            # Handle response shapes: either a direct list or a dict containing "Table" / "Table1"
            if isinstance(data, list):
                return data, "OK"
            elif isinstance(data, dict):
                for table_key in ["Table", "Table1", "data", "Data"]:
                    if table_key in data and isinstance(data[table_key], list):
                        return data[table_key], "OK"
                return [], "BSE response dict had no recognized table keys."
            else:
                return [], "BSE response was not a valid list or dictionary."

        except requests.RequestException as e:
            logger.warning(f"BSE network request error (attempt {attempt + 1}): {e}")
            time.sleep(backoff[attempt])
        except Exception as e:
            logger.warning(f"BSE fetch generic error: {e}")
            break

    return [], "BSE endpoint unreachable or returned invalid data after retries."


def normalize_bse_records(raw_items: List[Dict[str, Any]], deal_type_label: str) -> List[Dict[str, Any]]:
    """Standardizes raw BSE deals into the unified internal format."""
    normalized = []
    for item in raw_items:
        if not isinstance(item, dict):
            continue

        symbol = str(
            item.get("scrip_cd")
            or item.get("SCRIP_CD")
            or item.get("Scrip_CD")
            or item.get("scrip_name")
            or item.get("ScripName")
            or ""
        ).strip().upper()

        sec_name = str(
            item.get("scrip_name")
            or item.get("ScripName")
            or item.get("SCRIP_NAME")
            or symbol
        ).strip()

        client = str(
            item.get("client_name")
            or item.get("Client_Name")
            or item.get("CLIENT_NAME")
            or item.get("clientName")
            or ""
        ).strip()

        date_raw = (
            item.get("deal_date")
            or item.get("Deal_Date")
            or item.get("DEAL_DATE")
            or item.get("dt_tm")
            or item.get("DATE")
            or ""
        )
        deal_date = _parse_bse_date(date_raw)
        if not deal_date:
            continue

        deal_type_code = str(
            item.get("deal_type")
            or item.get("Deal_Type")
            or item.get("DEAL_TYPE")
            or ""
        ).strip().upper()

        # Buy / Sell parsing
        action_raw = str(
            item.get("deal_type")
            or item.get("Deal_Type")
            or item.get("trd_type")
            or item.get("TRD_TYPE")
            or item.get("buy_sell")
            or ""
        ).strip().upper()

        if action_raw in ["B", "BUY", "PURCHASE"]:
            buy_sell = "Buy"
        elif action_raw in ["S", "SELL", "SALE"]:
            buy_sell = "Sell"
        elif "BUY" in action_raw:
            buy_sell = "Buy"
        elif "SELL" in action_raw:
            buy_sell = "Sell"
        else:
            buy_sell = action_raw if action_raw else "Buy"

        qty = _parse_bse_number(item.get("qty") or item.get("Qty") or item.get("QTY") or item.get("Quantity"))
        price = _parse_bse_number(item.get("rate") or item.get("Rate") or item.get("RATE") or item.get("Price") or item.get("price"))

        if qty <= 0 or price <= 0:
            continue

        deal_val = qty * price
        deal_val_cr = deal_val / 10000000.0

        normalized.append({
            "exchange": "BSE",
            "deal_type": deal_type_label,
            "date": deal_date,
            "symbol": symbol,
            "security_name": sec_name if sec_name else symbol,
            "client_name": client if client else "NOT SPECIFIED",
            "buy_sell": buy_sell,
            "quantity": qty,
            "price": price,
            "deal_value": deal_val,
            "deal_value_cr": deal_val_cr,
        })

    return normalized


def fetch_bse_deals(from_date: date, to_date: date) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Fetches Bulk ('B') and Block ('K') deals from BSE for the given date span.
    """
    fdate_str = from_date.strftime("%Y%m%d")
    todate_str = to_date.strftime("%Y%m%d")

    status_info = {
        "status": "Loaded",
        "message": "BSE deals data loaded successfully.",
        "raw_count": 0,
        "cleaned_count": 0,
    }

    bulk_raw, _ = fetch_bse_endpoint("B", fdate_str, todate_str)
    block_raw, _ = fetch_bse_endpoint("K", fdate_str, todate_str)

    status_info["raw_count"] = len(bulk_raw) + len(block_raw)

    bulk_norm = normalize_bse_records(bulk_raw, "Bulk")
    block_norm = normalize_bse_records(block_raw, "Block")

    all_deals = bulk_norm + block_norm

    if not all_deals:
        status_info["status"] = "Warning"
        status_info["message"] = "No reported BSE Bulk/Block deals found for the requested period."
        return pd.DataFrame(), status_info

    df = pd.DataFrame(all_deals)
    df = df[(df["date"] >= from_date) & (df["date"] <= to_date)]
    status_info["cleaned_count"] = len(df)
    return df, status_info
