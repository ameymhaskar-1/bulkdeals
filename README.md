# NSE & BSE Deals Tracker
### Institutional Investment Banking Deal Flow & Historical Price Analytics Engine

A production-quality investment-banking tool for tracking, analyzing, and enriching Bulk and Block deals on the **National Stock Exchange of India (NSE)** and the **Bombay Stock Exchange (BSE)**.

---

## 🎯 Core Capabilities

1. **Exchange Ingestion with Fallbacks**:
   - **NSE**: Full session-cookie warm-up against `https://www.nseindia.com/` and report pages. Queries historical endpoints with automatic snapshot endpoint fallback.
   - **BSE**: Robust handling of `api.bseindia.com/BseIndiaAPI/api/BulkBlockDeal/w` supporting both direct array and nested table JSON schemas.
   - **Zero Crash Resilience**: Exponential backoff (1–5s) and safe JSON parsing that gracefully handles 403, 429, 500, HTML responses, and rate limits without throwing `JSONDecodeError`.

2. **Historical Market Price Context (T, T-1, T-5, T-15)**:
   - For every transaction, computes the **Deal Date Market Close (T)**, **1 Trading Day Ago Close (T-1)**, **5 Trading Days Ago Close (T-5)**, and **15 Trading Days Ago Close (T-15)**.
   - Accurately handles trading calendar logic (skipping weekends and market holidays).
   - Computes:
     - `Deal Price vs Deal Date Close (%)`
     - `1D Price Change %` (Deal Date Close vs T-1)
     - `5D Price Change %` (Deal Date Close vs T-5)
     - `15D Price Change %` (Deal Date Close vs T-15)
     - `Deal Price vs T-15 Close (%)`

3. **Executive Investment-Banking Reporting**:
   - Aggregate institutional inflow/outflow metrics (₹ Cr).
   - Concentration analysis: Top 15 Stocks, Top 15 Clients, and Repeat Institutional Activity.
   - Factual observations using institutional terminology without speculative bias.

4. **Multi-Sheet OpenPyXL Workbook Generation (`NSE_BSE_Deals_Tracker.xlsx`)**:
   - **`Raw Data`**: Excel Table `Deals_Data` with alternating stripes, freeze panes, auto-filters, and ₹ / Cr / % number formatting.
   - **`Analysis`**: Executive KPI cards, observations, Top 15 Stocks, Top 15 Clients, Deal Price Context, and Methodology.
   - **`Pivots`**: Structured cross-tabulation summaries (Exchange × Deal Type, Deal Type × Action, Daily Value, Symbol × Action).

---

## 🚀 Installation & Local Execution

### Prerequisites
- Python 3.9, 3.10, 3.11, or 3.12

### Step-by-Step Setup

1. **Clone the Repository**:
   ```bash
   git clone <repository_url>
   cd <repository_directory>
