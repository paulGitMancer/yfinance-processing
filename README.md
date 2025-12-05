# Stock Analysis Automation (Python)

This project automates the analysis of approximately 200 stock tickers using Python. 
It reads ticker symbols from an Excel file, retrieves market data from Yahoo Finance, 
performs calculations, and exports the results to a formatted Excel workbook.

## Features

- Reads stock tickers from an Excel file
- Retrieves current market data using yfinance
- Performs numerical calculations
- Writes results to a new Excel workbook
- Applies basic formatting to improve readability

## Input

- Excel file containing a list of ticker symbols
- The input file path is defined in the script (line 38)

## Output

- A formatted Excel workbook containing stock data and calculated values
- The output file path is defined in the script (line 165)

## Requirements

- Install dependencies:
<pre> pip install yfinance openpyxl</pre>

## Notes

- Designed for ~200 tickers but works with less or more
- Network performance may affect runtime
- Invalid tickers may result in missing values
