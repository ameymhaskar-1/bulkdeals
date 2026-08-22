"""
NSE Bulk and Block Deals Fetcher with Session Management and Fallbacks
"""

import time
import logging
from datetime import datetime, date
from typing import Dict, List, Tuple, Any, Optional
import requests
import pandas as pd

logger = logging.getLogger(__name__)

# ==============================================================================
# NSE UNOFFICIAL ENDPOINT DISCLAIMER:
# These are unofficial/reverse-engineered NSE endpoints, used in the same general
# manner as open-source libraries such as nsepython. They can change without notice.
# The application must fail gracefully and return an empty result plus a warning
# if NSE blocks the request or changes the endpoint/response structure.
# ==============================================================================

NSE_HOME_URL = "https://www.nseindia.com/"
NSE_REPORT_URL = "https://www.nseindia.com/report-detail/display-bulk-and-block-deals"
NSE_BULK_HIST_URL = "https://www.nseindia.com/api/historical/bulk-deals"
NSE_BLOCK_HIST_URL = "https://www.nseindia.com/api/historical/block-deals"
NSE_BULK_SNAP_URL = "https://www.nseindia.com/api/snapshot-capital-market-largedeal?index=bulk_deals"
NSE_BLOCK_SNAP_URL = "https://www.nseindia.com/api/snapshot-capital-market-largedeal?index=block_deals"

DEFAULT_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
}

API_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": NSE_REPORT_URL,
    "X-Requested-With": "XMLHttpRequest",
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin",
}


class NSESessionManager:
    """Manages an active HTTP session with NSE, warming up cookies and headers."""

    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update(DEFAULT_HEADERS)
        self.is_warmed_up = False

    def warm_up(self) -> bool:
        """Hits the homepage and deals report page to collect and retain required cookies."""
        try:
            r1 = self.session.get(NSE_HOME_URL, timeout=12)
            time.sleep(1.0)
            r2 = self.session.get(NSE_REPORT_URL, timeout=12)
            if r1.status_code in [200, 304] or r2.status_code in [200, 304]:
                self.is_warmed_up = True
                return True
        except Exception as exc:
            logger.warning(f"NSE Session warm-up warning: {exc}")
        return False

    def safe_get_json(self, url: str, params: Optional[Dict[str, Any]] = None, max_retries: int = 3) -> Tuple[Optional[Any], str]:
        """
        Executes an HTTP GET with retry, backoff, and robust JSON validation.
        NEVER blindly calls response.json().
        """
        if not self.is_warmed_up:
            self.warm_up()

        headers = API_HEADERS.copy()
        backoff_delays = [1.5, 3.5, 5.5]

        for attempt in range(max_retries):
            try:
                response = self.session.get(url, headers=headers, params=params, timeout=15)
                status = response.status_code

                if status == 403:
                    logger.warning(f"NSE returned 403 on attempt {attempt + 1}. Re-warming session...")
                    self.warm_up()
                    time.sleep(backoff_delays[min(attempt, len(backoff_delays) - 1)])
                    continue

                if status == 429:
                    logger.warning(f"NSE rate limited (429) on attempt {attempt + 1}. Backing off...")
                    time.sleep(backoff_delays[min(attempt, len(backoff_delays) - 1)] * 2)
                    continue

                if status in [500, 502, 503, 504]:
                    logger.warning(f"NSE server error {status} on attempt {attempt + 1}.")
                    time.sleep(backoff_delays[min(attempt, len(backoff_delays) - 1)])
                    continue

                if status != 200:
                    logger.warning(f"NSE request returned unexpected status {status}.")
                    time.sleep(backoff_delays[min(attempt, len(backoff_delays) - 1)])
                    continue

                # Content-Type and body check
                content_type = response.headers.get("Content-Type", "").lower()
                text = response.text.strip()

                if not text:
                    return None, "Empty response received from NSE."

                if text.startswith("<") or "html" in content_type:
                    # Received HTML instead of JSON
                    logger.warning("NSE returned HTML payload instead of expected JSON.")
                    time.sleep(backoff_delays[min(attempt, len(backoff_delays) - 1)])
                    continue

                try:
                    data = response.json()
                    return data, "OK"
                except Exception as json_err:
                    logger.warning(f"JSON decode failed on attempt {attempt + 1}: {json_err}")
                    time.sleep(backoff_delays[min(attempt, len(backoff_delays) - 1)])
                    continue

            except requests.RequestException as req_err:
                logger.warning(f"Network error during NSE fetch on attempt {attempt + 1}: {req_err}")
                time.sleep(backoff_delays[min(attempt, len(backoff_delays) - 1)])
            except Exception as e:
                logger.warning(f"Unexpected error during NSE fetch: {e}")
                break

        return None, "NSE API unreachable or blocked after retries."


def _parse_nse_number(val: Any) -> float:
    """Helper to convert string/numeric to float safely."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    try:
        s = str(val).replace(",", "").replace("₹", "").strip()
        return float(s) if s else 0.0
    except Exception:
        return 0.0


def _parse_nse_date(val: Any) -> Optional[date]:
    """Parses various NSE date formats to datetime.date object."""
    if not val:
        return None
    val_str = str(val).strip()
    for fmt in ("%d-%b-%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%b %d, %Y"):
        try:
            return datetime.strptime(val_str, fmt).date()
        except ValueError:
            continue
    return None


def normalize_nse_records(raw_items: List[Dict[str, Any]], deal_type_label: str) -> List[Dict[str, Any]]:
    """Normalizes raw NSE JSON items into the standard internal deal schema."""
    normalized = []
    for item in raw_items:
        if not isinstance(item, dict):
            continue

        # Symbol & Security
        symbol = str(
            item.get("symbol")
            or item.get("BD_SYMBOL")
            or item.get("symbolDesc")
            or item.get("secName")
            or item.get("securityName")
            or ""
        ).strip().upper()

        sec_name = str(
            item.get("securityName")
            or item.get("BD_SEC_NAME")
            or item.get("companyName")
            or symbol
        ).strip()

        # Client
        client = str(
            item.get("clientName")
            or item.get("BD_CLIENT_NAME")
            or item.get("client")
            or ""
        ).strip()

        # Date
        date_raw = (
            item.get("date")
            or item.get("BD_DT_DATE")
            or item.get("dealDate")
            or item.get("transDate")
            or ""
        )
        deal_date = _parse_nse_date(date_raw)
        if not deal_date:
            continue

        # Buy / Sell
        action_raw = str(
            item.get("buySell")
            or item.get("BD_BUY_SELL")
            or item.get("buy_sell")
            or item.get("action")
            or ""
        ).strip().upper()

        buy_sell = "Buy" if action_raw in ["BUY", "B", "PURCHASE"] else "Sell" if action_raw in ["SELL", "S", "SALE"] else action_raw

        # Quantity and Price
        quantity = _parse_nse_number(item.get("quantity") or item.get("BD_QTY_TRD") or item.get("qty") or item.get("tradedQty"))
        price = _parse_nse_number(item.get("tradePrice") or item.get("BD_TP_RATE") or item.get("price") or item.get("dealPrice") or item.get("avgPrice"))

        if quantity <= 0 or price <= 0:
            continue

        deal_val = quantity * price
        deal_val_cr = deal_val / 10000000.0

        normalized.append({
            "exchange": "NSE",
            "deal_type": deal_type_label,
            "date": deal_date,
            "symbol": symbol,
            "security_name": sec_name if sec_name else symbol,
            "client_name": client if client else "NOT SPECIFIED",
            "buy_sell": buy_sell,
            "quantity": quantity,
            "price": price,
            "deal_value": deal_val,
            "deal_value_cr": deal_val_cr,
        })

    return normalized


def fetch_nse_deals(from_date: date, to_date: date) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Fetches Bulk and Block deals from NSE across the specified date window.
    Returns a normalized DataFrame and a source status diagnostic dict.
    """
    mgr = NSESessionManager()
    mgr.warm_up()

    from_str = from_date.strftime("%d-%m-%Y")
    to_str = to_date.strftime("%d-%m-%Y")

    status_info = {
        "status": "Loaded",
        "message": "NSE historical deal data loaded successfully.",
        "used_snapshot_fallback": False,
        "raw_count": 0,
        "cleaned_count": 0,
    }

    all_normalized_records: List[Dict[str, Any]] = []

    # 1. Try Historical Bulk Deals
    bulk_data, bulk_msg = mgr.safe_get_json(NSE_BULK_HIST_URL, params={"from": from_str, "to": to_str})
    bulk_items = []
    if isinstance(bulk_data, dict) and "data" in bulk_data and isinstance(bulk_data["data"], list):
        bulk_items = bulk_data["data"]
    elif isinstance(bulk_data, list):
        bulk_items = bulk_data

    # 2. Try Historical Block Deals
    block_data, block_msg = mgr.safe_get_json(NSE_BLOCK_HIST_URL, params={"from": from_str, "to": to_str})
    block_items = []
    if isinstance(block_data, dict) and "data" in block_data and isinstance(block_data["data"], list):
        block_items = block_data["data"]
    elif isinstance(block_data, list):
        block_items = block_data

    status_info["raw_count"] = len(bulk_items) + len(block_items)

    # 3. Snapshot Fallback if historical is blocked / empty
    if len(bulk_items) == 0 and len(block_items) == 0:
        logger.info("NSE historical returned 0 records. Trying snapshot endpoint fallback...")
        snap_bulk, _ = mgr.safe_get_json(NSE_BULK_SNAP_URL)
        snap_block, _ = mgr.safe_get_json(NSE_BLOCK_SNAP_URL)

        if isinstance(snap_bulk, dict) and "data" in snap_bulk and isinstance(snap_bulk["data"], list):
            bulk_items = snap_bulk["data"]
        elif isinstance(snap_bulk, list):
            bulk_items = snap_bulk

        if isinstance(snap_block, dict) and "data" in snap_block and isinstance(snap_block["data"], list):
            block_items = snap_block["data"]
        elif isinstance(snap_block, list):
            block_items = snap_block

        if len(bulk_items) > 0 or len(block_items) > 0:
            status_info["used_snapshot_fallback"] = True
            status_info["status"] = "Warning"
            status_info["message"] = "NSE historical endpoint unavailable — snapshot fallback used."

    # Normalize records
    bulk_norm = normalize_nse_records(bulk_items, "Bulk")
    block_norm = normalize_nse_records(block_items, "Block")
    all_normalized_records.extend(bulk_norm)
    all_normalized_records.extend(block_norm)

    if not all_normalized_records:
        if status_info["status"] == "Loaded":
            status_info["status"] = "Warning"
            status_info["message"] = "No reported NSE Bulk/Block deals found for the requested period."
        else:
            status_info["status"] = "Failed"
            status_info["message"] = "NSE endpoints blocked the connection or returned invalid payloads."
        return pd.DataFrame(), status_info

    df = pd.DataFrame(all_normalized_records)
    # Filter by user-selected date range (important for snapshot fallbacks)
    df = df[(df["date"] >= from_date) & (df["date"] <= to_date)]

    status_info["cleaned_count"] = len(df)
    return df, status_info
