# Portfolio Rebalancer App

A Streamlit-based portfolio rebalancing tool that helps users shift their investment portfolio to match predefined conservative, medium, or aggressive strategies.

## Features

✅ **Three Portfolio Strategies:**
- Conservative: 70% bonds, 30% stocks
- Medium: 60% stocks, 40% bonds
- Aggressive: 80% stocks, 20% bonds

✅ **Flexible Input:**
- Manual entry of holdings (ticker, shares, price)
- CSV upload support

✅ **Dashboard Visualization:**
- Current portfolio pie charts and breakdown tables
- Recommended strategy visualization
- Side-by-side comparison with arrow flow

✅ **Excel Export:**
- 3-sheet Excel report with current holdings, recommended allocation, and trade instructions
- All calculations and dollar amounts included
- Professional formatting

## Setup Instructions

### 1. Install Python & Anaconda
Make sure you have Anaconda installed. If not, download from: https://www.anaconda.com/download

### 2. Create Environment (Optional but Recommended)
```bash
conda create -n portfolio_rebalancer python=3.10
conda activate portfolio_rebalancer
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

Or install individually:
```bash
pip install streamlit pandas yfinance openpyxl matplotlib
```

## Running the App

### From Terminal/Command Line:
```bash
cd /path/to/portfolio_rebalancer
streamlit run portfolio_rebalancer.py
```

The app will open in your browser at http://localhost:8501

### From VS Code:
1. Open the project folder in VS Code
2. Open terminal (Ctrl+` or Cmd+`)
3. Run: streamlit run portfolio_rebalancer.py

## How to Use

1. **Select Risk Tolerance** (left sidebar)
   - Conservative, Medium, or Aggressive

2. **Select International Preference**
   - US only or include international stocks

3. **Input Your Portfolio**
   - Option A: Manual entry - enter each ticker, shares, and current price
   - Option B: CSV upload - upload a file with columns: Ticker, Shares, Price

4. **Review Dashboard**
   - See current allocation pie chart and breakdown
   - View recommended strategy with target allocation
   - See specific buy/sell recommendations

5. **Export Report**
   - Click "Generate Excel Report"
   - Download Excel file with 3 sheets:
     - Current Portfolio (your holdings breakdown)
     - Recommended Portfolio (target allocation)
     - Trading Instructions (specific buy/sell actions)

## Sample Portfolio CSV Format

```
Ticker,Shares,Price
AAPL,10,195.50
TSLA,5,250.00
VOO,20,450.00
BND,30,78.50
```

## File Structure

```
portfolio_rebalancer/
├── portfolio_rebalancer.py      # Main app
├── requirements.txt              # Dependencies
├── sample_portfolio.csv          # Sample data for testing
└── README.md                     # This file
```

---

Good luck with your FPD project! 🚀
