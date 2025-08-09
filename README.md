# DCF & Comps Valuation Tool

A Streamlit-based valuation model builder for NSE-listed companies.  
This tool automates peer selection, market data retrieval, fundamental scraping, forecasting of financials, and visualization via Football Field Analysis — all in one workflow.

---

## Features

### 1. **Company & Peer Selection**
- User enters an NSE ticker (without `.NS`, e.g., `TCS`).
- Peers are selected from `EQUITY_Final.xlsx` using:
  1. Macro → Sector → Industry → Basic Industry filters
  2. If fewer than **10 peers** match, filters are relaxed step-by-step until the threshold is met.
- **Price history is downloaded first** — peers without sufficient weekly data are removed before processing.

---

### 2. **Data Retrieval**
#### **Price Data**
- 5 years of **weekly prices** for NIFTY50 + peers.
- Used for:
  - Beta calculation (Regression Beta & Bottom-up Beta)
  - Trading Comps multiples
  - Football Field charts

#### **Fundamental Data (Yahoo Finance)**
- **Market Metrics:** Market Cap, Shares Outstanding, Current Price, Total Debt, Diluted EPS
- **Income Statement:** Total Revenue, EBIT, EBITDA
- **Balance Sheet Items:**
  - Cash & Cash Equivalents
  - Restricted Cash
  - Other Short Term Investments
  - Investment in Financial Assets
  - Available For Sale Securities
  - Other Investments
  - Goodwill
  - Other Intangible Assets
  - Deferred Tax Assets
  - Non-Current Deferred Assets
  - Fixed Asset Revaluation Reserve
  - Stockholders' Equity
  - Minority Interest

#### **Financial Statements (Screener.in)**
- Profit & Loss Statement
- Balance Sheet

---

### 3. **Forecasting Methods**
For each P/L and Balance Sheet item, the tool can forecast using:
1. **Constant Growth Method**
2. **Average Year-on-Year (YoY) Growth**
3. **CAGR**
4. **Linear Regression**
5. **Trend Method**
6. **Exponential Smoothing (ETS)**
7. **Simple Moving Average (SMA, 3-year)**
8. **Weighted Moving Average (WMA, 3-year)**
9. **% of Revenues**

---

### 4. **Valuation Models**
#### **Trading Comparables**
- Calculates multiples such as:
  - EV/EBITDA
  - EV/EBIT
  - EV/Sales
  - P/E
- Automatically computed for peers.

#### **DCF Valuation**
- Forecasted cash flows discounted to present value.
- Terminal value calculated using:
  1. **Exit Multiple Method**
  2. **Gordon Growth Model**
- Growth rate (`g`) estimated using both methods for comparison.

---

### 5. **Football Field Analysis**
- Combines:
  - DCF Valuation range (Equity Value & Enterprise Value)
  - Trading Comparables multiples range
- Draws ranges for both **Equity Value** and **Enterprise Value**.
- Median/midpoints highlighted for easy interpretation.

---

### 6. **Beta Calculation**
- **Regression Beta:** From 5 years of weekly returns vs. NIFTY50.
- **Bottom-Up Beta:** Industry average adjusted for leverage.
- User can select which beta to apply in the model.

---

##  Output
The tool generates an Excel workbook `<TICKER>_DCF_Comps_Valuation_Tool.xlsx` containing:
- `PEERS` — Comparable companies’ fundamentals.
- `RAW_P_L` — Profit & Loss statement from Screener.
- `RAW_B_S` — Balance Sheet from Screener.
- `PRICE_HISTORY` — Weekly stock prices for NIFTY50 + peers.
- Valuation outputs and Football Field Analysis integrated into `DCF_Template.xlsx`.

---

##  How to Run
```bash
# Install dependencies
pip install -r requirements.txt

# Run Streamlit app
streamlit run "DCF model.py"
