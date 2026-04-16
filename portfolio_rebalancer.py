import streamlit as st
import pandas as pd
import yfinance as yf
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.chart import PieChart, Reference
import io
import os
from datetime import datetime

# Page config
st.set_page_config(page_title="Portfolio Rebalancer", layout="wide")
st.title("📊 Portfolio Rebalancer")

# Define strategies
STRATEGIES = {
    "Conservative": {
        "bonds": 0.70,
        "stocks": 0.30,
        "us_stocks": 0.70,
        "intl_stocks": 0.30
    },
    "Medium": {
        "bonds": 0.40,
        "stocks": 0.60,
        "us_stocks": 0.70,
        "intl_stocks": 0.30
    },
    "Aggressive": {
        "bonds": 0.20,
        "stocks": 0.80,
        "us_stocks": 0.70,
        "intl_stocks": 0.30
    }
}

# Sidebar for input
st.sidebar.header("Portfolio Setup")

# Step 1: Risk tolerance
risk_level = st.sidebar.radio(
    "What's your risk tolerance?",
    ["Conservative", "Medium", "Aggressive"],
    help="Conservative = focus on stability, Aggressive = focus on growth"
)

# Step 2: International interest
intl_interest = st.sidebar.radio(
    "Interest in international stocks?",
    ["No, US only", "Yes, include international"],
    help="International diversifies your portfolio geographically"
)

# Adjust international split based on preference
if intl_interest == "No, US only":
    STRATEGIES[risk_level]["us_stocks"] = 1.0
    STRATEGIES[risk_level]["intl_stocks"] = 0.0
else:
    STRATEGIES[risk_level]["us_stocks"] = 0.70
    STRATEGIES[risk_level]["intl_stocks"] = 0.30

st.sidebar.markdown("---")
st.sidebar.header("Portfolio Input")

# Portfolio input method
input_method = st.sidebar.radio(
    "How do you want to input your portfolio?",
    ["Manual Entry", "CSV Upload"],
    help="Enter holdings manually or upload a CSV file"
)

portfolio_data = {}

if input_method == "Manual Entry":
    st.sidebar.markdown("**Enter your holdings:**")
    num_holdings = st.sidebar.number_input("How many holdings?", min_value=1, max_value=20, value=3)

    for i in range(num_holdings):
        st.sidebar.markdown(f"**Holding {i+1}**")
        company_search = st.sidebar.text_input(f"Company name or ticker", key=f"search_{i}", placeholder="e.g. Apple or AAPL")

        ticker = ""
        price = 0.0

        if company_search:
            try:
                results = yf.Search(company_search, max_results=5).quotes
                if results:
                    options = {
                        f"{r.get('shortname', r.get('symbol', ''))} ({r['symbol']})": r['symbol']
                        for r in results if 'symbol' in r
                    }
                    selected = st.sidebar.selectbox(f"Select company", list(options.keys()), key=f"select_{i}")
                    ticker = options[selected]
                    try:
                        price = round(yf.Ticker(ticker).fast_info.last_price, 2)
                        st.sidebar.success(f"Live price: ${price:.2f}")
                    except:
                        price = st.sidebar.number_input(f"Price (auto-fetch failed)", min_value=0.0, key=f"price_{i}")
                else:
                    ticker = company_search.upper()
                    price = st.sidebar.number_input(f"Price", min_value=0.0, key=f"price_{i}")
            except:
                ticker = company_search.upper()
                price = st.sidebar.number_input(f"Price", min_value=0.0, key=f"price_{i}")

        shares = st.sidebar.number_input(f"Shares", min_value=0.0, key=f"shares_{i}")

        if ticker and shares > 0 and price > 0:
            portfolio_data[ticker] = {"shares": shares, "price": price}

elif input_method == "CSV Upload":
    uploaded_file = st.sidebar.file_uploader("Upload CSV (Ticker, Shares, Price)", type="csv")
    
    if uploaded_file:
        df = pd.read_csv(uploaded_file)
        # Handle different column names
        ticker_col = next((col for col in df.columns if col.upper() in ["TICKER", "SYMBOL", "STOCK"]), None)
        shares_col = next((col for col in df.columns if col.upper() in ["SHARES", "QUANTITY"]), None)
        price_col = next((col for col in df.columns if col.upper() in ["PRICE", "CURRENT PRICE", "COST"]), None)
        
        if ticker_col and shares_col and price_col:
            for _, row in df.iterrows():
                ticker = str(row[ticker_col]).upper().strip()
                shares = float(row[shares_col])
                price = float(row[price_col])
                if ticker and shares > 0:
                    portfolio_data[ticker] = {"shares": shares, "price": price}
        else:
            st.sidebar.error("CSV must have Ticker, Shares, and Price columns")

# Calculate current portfolio
if portfolio_data:
    # Calculate total value and current holdings
    total_value = sum(data["shares"] * data["price"] for data in portfolio_data.values())
    
    current_allocation = {}
    for ticker, data in portfolio_data.items():
        holding_value = data["shares"] * data["price"]
        current_allocation[ticker] = {
            "shares": data["shares"],
            "price": data["price"],
            "value": holding_value,
            "percent": (holding_value / total_value) * 100
        }
    
    # Display current portfolio
    st.header("📈 Current Portfolio")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Portfolio Value", f"${total_value:,.2f}")
    with col2:
        st.metric("Number of Holdings", len(portfolio_data))
    with col3:
        st.metric("Current Strategy", risk_level)
    
    # Current portfolio pie chart
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Current Holdings Breakdown")
        current_df = pd.DataFrame([
            {
                "Ticker": ticker,
                "Shares": data["shares"],
                "Price": f"${data['price']:.2f}",
                "Value": f"${data['value']:,.2f}",
                "Allocation %": f"{data['percent']:.1f}%"
            }
            for ticker, data in current_allocation.items()
        ])
        st.dataframe(current_df, use_container_width=True, hide_index=True)
        
        # Create pie chart for current
        import matplotlib.pyplot as plt
        fig, ax = plt.subplots(figsize=(6, 4))
        values = [data["percent"] for data in current_allocation.values()]
        labels = [f"{ticker}\n{data['percent']:.1f}%" for ticker, data in current_allocation.items()]
        ax.pie(values, labels=labels, autopct='', startangle=90)
        ax.set_title("Current Portfolio Allocation", fontsize=12, fontweight='bold')
        st.pyplot(fig)
    
    # Recommended strategy
    with col2:
        st.subheader(f"Recommended {risk_level} Strategy")
        strategy = STRATEGIES[risk_level]
        
        strategy_text = f"""
        **Asset Allocation Target:**
        - **Stocks:** {strategy['stocks']*100:.0f}%
        - **Bonds:** {strategy['bonds']*100:.0f}%
        
        **Stock Breakdown:**
        - **US Stocks:** {strategy['us_stocks']*100:.0f}%
        - **International:** {strategy['intl_stocks']*100:.0f}%
        """
        st.markdown(strategy_text)
        
        # Recommended allocation visualization
        fig2, ax2 = plt.subplots(figsize=(6, 4))
        recommended_labels = [
            f"Stocks\n{strategy['stocks']*100:.0f}%",
            f"Bonds\n{strategy['bonds']*100:.0f}%"
        ]
        recommended_values = [strategy['stocks'], strategy['bonds']]
        ax2.pie(recommended_values, labels=recommended_labels, autopct='', startangle=90, colors=['#2E86AB', '#A23B72'])
        ax2.set_title(f"Recommended {risk_level} Allocation", fontsize=12, fontweight='bold')
        st.pyplot(fig2)
    
    # Arrow and transition
    st.markdown("---")
    st.markdown("<h3 style='text-align: center;'>⬇️ Rebalancing Plan ⬇️</h3>", unsafe_allow_html=True)
    st.markdown("---")
    
    # Calculate recommended holdings
    target_stocks_value = total_value * strategy['stocks']
    target_bonds_value = total_value * strategy['bonds']
    
    # Categorize current holdings
    stock_holdings = {}
    bond_holdings = {}
    uncategorized = {}
    
    # Common ETFs and their categories
    stock_etfs = ['VOO', 'VTI', 'SPY', 'QQQ', 'IVV', 'VEA', 'VXUS', 'EEM', 'VWO']
    bond_etfs = ['BND', 'BLV', 'AGG', 'IEF', 'TLT', 'SHV', 'LQD']
    
    for ticker in current_allocation.keys():
        if ticker in stock_etfs:
            stock_holdings[ticker] = current_allocation[ticker]['value']
        elif ticker in bond_etfs:
            bond_holdings[ticker] = current_allocation[ticker]['value']
        else:
            # Default: assume stocks (users can override)
            stock_holdings[ticker] = current_allocation[ticker]['value']
    
    current_stock_value = sum(stock_holdings.values())
    current_bond_value = sum(bond_holdings.values())
    
    # Calculate trades
    stock_diff = target_stocks_value - current_stock_value
    bond_diff = target_bonds_value - current_bond_value
    
    st.subheader("Recommended Trades to Reach Target")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**Stock Position**")
        st.metric("Current", f"${current_stock_value:,.2f}", f"{(current_stock_value/total_value)*100:.1f}%")
        st.metric("Target", f"${target_stocks_value:,.2f}", f"{strategy['stocks']*100:.0f}%")
        if stock_diff > 0:
            st.success(f"🟢 BUY ${abs(stock_diff):,.2f} in stock ETFs (e.g., VOO)")
        elif stock_diff < 0:
            st.error(f"🔴 SELL ${abs(stock_diff):,.2f} of stock holdings")
        else:
            st.info("✅ No action needed")
    
    with col2:
        st.markdown("**Bond Position**")
        st.metric("Current", f"${current_bond_value:,.2f}", f"{(current_bond_value/total_value)*100:.1f}%")
        st.metric("Target", f"${target_bonds_value:,.2f}", f"{strategy['bonds']*100:.0f}%")
        if bond_diff > 0:
            st.success(f"🟢 BUY ${abs(bond_diff):,.2f} in bonds (e.g., BND)")
        elif bond_diff < 0:
            st.error(f"🔴 SELL ${abs(bond_diff):,.2f} of bond holdings")
        else:
            st.info("✅ No action needed")
    
    st.markdown("---")
    
    # Export to Excel
    st.subheader("📥 Export Report")
    
    if st.button("Generate Excel Report", key="export_btn"):
        # Create workbook
        wb = Workbook()
        wb.remove(wb.active)  # Remove default sheet
        
        # Sheet 1: Current Portfolio
        ws1 = wb.create_sheet("Current Portfolio")
        ws1['A1'] = "CURRENT PORTFOLIO"
        ws1['A1'].font = ws1['A1'].font.copy()
        ws1['A1'].font = ws1['A1'].font.copy()
        
        ws1['A3'] = "Ticker"
        ws1['B3'] = "Shares"
        ws1['C3'] = "Price"
        ws1['D3'] = "Value"
        ws1['E3'] = "Allocation %"
        
        for idx, (ticker, data) in enumerate(current_allocation.items(), start=4):
            ws1[f'A{idx}'] = ticker
            ws1[f'B{idx}'] = data['shares']
            ws1[f'C{idx}'] = data['price']
            ws1[f'D{idx}'] = data['value']
            ws1[f'E{idx}'] = data['percent']
        
        # Add totals
        last_row = 4 + len(current_allocation)
        ws1[f'A{last_row}'] = "TOTAL"
        ws1[f'D{last_row}'] = total_value
        
        ws1.column_dimensions['A'].width = 12
        ws1.column_dimensions['B'].width = 12
        ws1.column_dimensions['C'].width = 12
        ws1.column_dimensions['D'].width = 15
        ws1.column_dimensions['E'].width = 15
        
        # Sheet 2: Recommended Portfolio
        ws2 = wb.create_sheet("Recommended Portfolio")
        ws2['A1'] = f"RECOMMENDED {risk_level.upper()} PORTFOLIO"
        
        ws2['A3'] = "Asset Class"
        ws2['B3'] = "Target %"
        ws2['C3'] = "Target Value"
        ws2['D3'] = "Current Value"
        ws2['E3'] = "Difference"
        
        ws2['A4'] = "Stocks"
        ws2['B4'] = strategy['stocks']
        ws2['C4'] = target_stocks_value
        ws2['D4'] = current_stock_value
        ws2['E4'] = stock_diff
        
        ws2['A5'] = "Bonds"
        ws2['B5'] = strategy['bonds']
        ws2['C5'] = target_bonds_value
        ws2['D5'] = current_bond_value
        ws2['E5'] = bond_diff
        
        ws2['A7'] = "Stock Breakdown:"
        ws2['A8'] = "US Stocks"
        ws2['B8'] = strategy['us_stocks']
        ws2['A9'] = "International"
        ws2['B9'] = strategy['intl_stocks']
        
        ws2.column_dimensions['A'].width = 20
        ws2.column_dimensions['B'].width = 15
        ws2.column_dimensions['C'].width = 15
        ws2.column_dimensions['D'].width = 15
        ws2.column_dimensions['E'].width = 15
        
        # Sheet 3: Trading Instructions
        ws3 = wb.create_sheet("Trading Instructions")
        ws3['A1'] = "REBALANCING TRADES"
        
        ws3['A3'] = "Action"
        ws3['B3'] = "Asset Class"
        ws3['C3'] = "Amount"
        ws3['D3'] = "Suggested Holdings"
        
        row = 4
        if stock_diff > 0:
            ws3[f'A{row}'] = "BUY"
            ws3[f'B{row}'] = "Stocks"
            ws3[f'C{row}'] = stock_diff
            ws3[f'D{row}'] = "VOO, VTI, or similar broad market ETFs"
            row += 1
        elif stock_diff < 0:
            ws3[f'A{row}'] = "SELL"
            ws3[f'B{row}'] = "Stocks"
            ws3[f'C{row}'] = abs(stock_diff)
            ws3[f'D{row}'] = "Reduce highest allocation holdings"
            row += 1
        
        if bond_diff > 0:
            ws3[f'A{row}'] = "BUY"
            ws3[f'B{row}'] = "Bonds"
            ws3[f'C{row}'] = bond_diff
            ws3[f'D{row}'] = "BND, AGG, or similar bond ETFs"
            row += 1
        elif bond_diff < 0:
            ws3[f'A{row}'] = "SELL"
            ws3[f'B{row}'] = "Bonds"
            ws3[f'C{row}'] = abs(bond_diff)
            ws3[f'D{row}'] = "Reduce bond holdings proportionally"
            row += 1
        
        ws3['A6'] = "Summary:"
        ws3['A7'] = f"Total Portfolio Value: ${total_value:,.2f}"
        ws3['A8'] = f"Strategy: {risk_level}"
        ws3['A9'] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        
        ws3.column_dimensions['A'].width = 12
        ws3.column_dimensions['B'].width = 15
        ws3.column_dimensions['C'].width = 15
        ws3.column_dimensions['D'].width = 30
        
        # Save to bytes
        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)
        
        # Download button
        st.download_button(
            label="📥 Download Excel Report",
            data=excel_buffer.getvalue(),
            file_name=f"Portfolio_Rebalance_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        st.success("✅ Excel report generated successfully!")

else:
    st.info("👈 Enter your portfolio holdings on the left to get started!")
