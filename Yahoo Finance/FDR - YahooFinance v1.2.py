import subprocess
import os
import yfinance as yf
import pandas as pd

# Ensure xlsxwriter is installed
subprocess.call(['pip', 'install', 'xlsxwriter'])

# Define the ticker symbols you want to analyze
ticker_symbols = ['AAPL', 'MSFT']  # Example list of ticker symbols

# Initialize a dictionary to store the latest date of each ticker
latest_dates = {}

# Fetch data for each ticker
for ticker_symbol in ticker_symbols:
    # Fetch the financial statements
    ticker = yf.Ticker(ticker_symbol)
    income_statement = ticker.financials
    balance_sheet = ticker.balance_sheet
    cash_flow = ticker.cashflow
    historical_data = ticker.history(period="max")  # Fetch historical daily data

    # Convert the datetime index to timezone-unaware datetime index
    historical_data.index = historical_data.index.tz_localize(None)

    # Save the latest date of the historical data for this ticker
    latest_dates[ticker_symbol] = historical_data.index.max()

    # Step 3: Create an Excel writer using pandas
    file_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads', f"{ticker_symbol}_YF_retrieved.xlsx")

    # Save each statement in a different sheet
    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:

        # Save general information in a single sheet
        general_info_df = pd.DataFrame({
            'Company Name': [ticker.info.get('longName', 'N/A')],
            'Ticker': [ticker_symbol],
            'Currency': [ticker.info.get('currency', 'N/A')],
            'Exchange Market': [ticker.info.get('exchange', 'N/A')],
            'Date': [latest_dates[ticker_symbol].strftime('%Y-%m-%d')],
            'Previous Close': [ticker.info.get('previousClose', 'N/A')],
            'Open': [ticker.info.get('open', 'N/A')],
            'Market Cap': [ticker.info.get('marketCap', 'N/A')],
            'Beta': [ticker.info.get('beta', 'N/A')],
            'PE Ratio': [ticker.info.get('trailingPE', 'N/A')],
            'EPS': [ticker.info.get('eps', 'N/A')],
            'Dividend Yield': [ticker.info.get('dividendYield', 'N/A')],
            'Debt to Equity': [ticker.info.get('debtToEquity', 'N/A')],
            'Gross Profit Margin': [ticker.info.get('grossMargins', 'N/A')],
            'Current Ratio': [ticker.info.get('currentRatio', 'N/A')],
            'Return on Assets (ROA)': [ticker.info.get('returnOnAssets', 'N/A')],
            'Return on Equity (ROE)': [ticker.info.get('returnOnEquity', 'N/A')],
            'Consensus EPS Estimate': [ticker.info.get('forwardEps', 'N/A')]
        })

        general_info_df.to_excel(writer, sheet_name=f'{ticker_symbol} - Information', index=False)

        # Save historical data on a daily frequency over the whole available period
        historical_data.to_excel(writer, sheet_name='Historical Data - All', index=True)

        # Filter data for the last year
        last_year_data = historical_data.loc[pd.Timestamp.now() - pd.DateOffset(years=1):]

        # Save daily data for the last year
        last_year_data.to_excel(writer, sheet_name='Historical Data - Last Year', index=True)

        # Save financial statements
        income_statement_flipped = income_statement.iloc[::-1]  # Flip rows
        balance_sheet_flipped = balance_sheet.iloc[::-1]  # Flip rows
        cash_flow_flipped = cash_flow.iloc[::-1]  # Flip rows

        # Save financial statements
        income_statement_flipped.to_excel(writer, sheet_name='Income Statement')
        balance_sheet_flipped.to_excel(writer, sheet_name='Balance Sheet')
        cash_flow_flipped.to_excel(writer, sheet_name='Cash Flow')

        print("Financial data for", ticker_symbol, "saved successfully to", file_path)
