# Cryptocurrency Data Fetcher and Excel Updater

This project fetches live cryptocurrency data from the CoinMarketCap API, processes it, and updates an Excel file with the latest information. The data includes cryptocurrency name, symbol, price, market cap, 24-hour volume, and 24-hour change percentage. The project also performs analysis on the top 5 cryptocurrencies by market cap and updates the Excel sheet with relevant details.

## Features
- Fetches live cryptocurrency data from CoinMarketCap.
- Processes and organizes cryptocurrency data.
- Updates the Excel file with the latest data.
- Performs data analysis to identify the top 5 cryptocurrencies by market cap.
- Applies conditional formatting (color-coding) to the 24-hour change percentage and market cap columns in Excel.
- Saves the Excel file with the updated data.

## Installation

To get started with the project, you need to install the required dependencies. Follow these steps:

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/yourrepository.git
   cd yourrepository
Activate the virtual environment:

For Windows:
```bash
venv\Scripts\activate
For macOS/Linux:
```bash
source venv/bin/activate

Install the required dependencies:

```bash
pip install -r requirements.txt

Usage
Make sure you have a valid CoinMarketCap API key. You can get it from CoinMarketCap API.

Run the script:

bash
Copy code
python yourscript.py
When prompted, enter the name of the Excel file where you want the cryptocurrency data to be updated. The script will fetch live data and update the Excel file every 10 seconds (or at the defined interval).

The Excel file will be updated with the latest cryptocurrency data and analysis, including conditional formatting.

File Format
The Excel file will have the following columns:

Name: Cryptocurrency name (e.g., Bitcoin)
Symbol: Cryptocurrency symbol (e.g., BTC)
Price (USD): The current price in USD
Market Cap: The market capitalization of the cryptocurrency
24h Volume: The 24-hour trading volume
24h Change (%): The 24-hour percentage change in price
The script will also update the top 5 cryptocurrencies by market cap and apply conditional coloring based on the 24-hour change percentage.

License
This project is licensed under the MIT License - see the LICENSE file for details.
