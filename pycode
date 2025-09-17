import yfinance as yf
import pandas as pd
from yahooquery import search
from datetime import datetime

def get_stock_symbol_online(company_name: str) -> str:
    """
    Automatically search Yahoo Finance for a company name
    and return the stock symbol.
    """
    results = search(company_name)
    if results and "quotes" in results and len(results["quotes"]) > 0:
        symbol = results["quotes"][0]["symbol"]
        return symbol
    else:
        raise ValueError(f"No stock symbol found for '{company_name}'")

def fetch_and_save_stock_data(company_name: str, start_date_str: str, end_date_str: str):
    """
    Fetch stock data for a company between start_date and end_date (DD-MM-YYYY),
    remove holidays (Volume = 0), keep only Open, High, Low, Close, Volume,
    and save to Excel with proper file name.
    """
    # Convert input dates from DD-MM-YYYY to YYYY-MM-DD for yfinance
    start_date = datetime.strptime(start_date_str, "%d-%m-%Y").strftime("%Y-%m-%d")
    end_date = datetime.strptime(end_date_str, "%d-%m-%Y").strftime("%Y-%m-%d")
    
    symbol = get_stock_symbol_online(company_name)
    
    # Download stock data
    data = yf.download(symbol, start=start_date, end=end_date, interval="1d", auto_adjust=False)
    
    # Reset index so Date becomes a column
    data.reset_index(inplace=True)
    
    # Remove holidays / non-trading days (Volume = 0)
    #data = data[data["Volume"] > 0]
    
    # Flatten MultiIndex columns if any
    if isinstance(data.columns, pd.MultiIndex):
        data.columns = [col[0] for col in data.columns]
    
    # Remove 'Adj Close' column if exists
    if 'Adj Close' in data.columns:
        data = data.drop(columns=['Adj Close'])
    
    # Keep only desired columns in order
    data = data[['Date', 'Open', 'High', 'Low', 'Close', 'Volume']]
    
    # Convert Date column to DD-MM-YYYY format for Excel
    data['Date'] = data['Date'].dt.strftime('%d-%m-%Y')
    
    # Generate Excel file name like: "COMPANYNAME_01-07 to 16-07.xlsx"
    start_str = datetime.strptime(start_date_str, "%d-%m-%Y").strftime("%d-%m")
    end_str = datetime.strptime(end_date_str, "%d-%m-%Y").strftime("%d-%m")
    filename = f"{company_name.replace(' ', '_')}_{start_str}_to_{end_str}.xlsx"
    
    # Save to Excel
    data.to_excel(filename, index=False)
    
    print(f"Cleaned data for '{company_name}' ({symbol}) saved to {filename}")

# ------------------- USAGE -------------------
if __name__ == "__main__":
    company_name = input("Enter company full name or keyword: ")
    start_date_str = input("Enter start date (DD-MM-YYYY): ")
    end_date_str = input("Enter end date (DD-MM-YYYY): ")
    
    try:
        fetch_and_save_stock_data(company_name, start_date_str, end_date_str)
    except Exception as e:
        print(e)
