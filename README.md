# Stock Data Downloader

A Python script to fetch stock market data for any company using Yahoo Finance API, clean it, and save to Excel.

## Features
- Search company by full name or short form (e.g., "Tata Steel" or "TISCO").
- Download stock data between given start and end dates.
- Remove holidays and non-trading days.
- Export clean Excel file with columns: Date, Open, High, Low, Close, Volume.
- File automatically named as: `CompanyName_01-07_to_16-07.xlsx`.

## Requirements
- Python 3.10+
- Install dependencies:
  ```bash

  pip install yfinance pandas yahooquery openpyxl

