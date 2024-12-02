# Cryptocurrency Data Fetcher and Excel Updater

This project fetches live cryptocurrency data from the CoinMarketCap API, processes it, and updates an Excel file with the latest information. The data includes cryptocurrency name, symbol, price, market cap, 24-hour volume, and 24-hour percentage change. Additionally, it identifies the top 5 cryptocurrencies by market cap, performs analysis, and applies conditional formatting in the Excel file.

## Features

- Fetches live cryptocurrency data from CoinMarketCap.
- Processes and organizes cryptocurrency data into a user-friendly format.
- Updates an Excel file with the following columns:
  - **Name**: Cryptocurrency name (e.g., Bitcoin)
  - **Symbol**: Cryptocurrency symbol (e.g., BTC)
  - **Price (USD)**: The current price in USD
  - **Market Cap**: The market capitalization of the cryptocurrency
  - **24h Volume**: The 24-hour trading volume
  - **24h Change (%)**: The 24-hour percentage change in price
- Identifies and highlights the top 5 cryptocurrencies by market cap.
- Applies conditional formatting to:
  - **24h Change (%)**: Color-coded to indicate positive or negative changes.
  - **Market Cap**: Color-coded based on the value.
- Saves the Excel file with updated data and formatting.

## Installation

To get started with the project, you need to install the required dependencies. Follow these steps:

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/yourrepository.git
   cd yourrepository
2. Activate the virtual environment:
   For Windows:
   ```bash
   venv\Scripts\activate
3. For macOS/Linux:
   ```bash
   source venv/bin/activate

4. Install the required dependencies:
   ```bash
   pip install -r requirements.txt

**Usage**
Make sure you have a valid CoinMarketCap API key. You can get it from CoinMarketCap API.

5. Run the script:
   ```bash
   python yourscript.py

When prompted, enter the name of the Excel file where you want the cryptocurrency data to be updated. The script will fetch live data and update the Excel file every 10 seconds (or at the defined interval).

The Excel file will be updated with the latest cryptocurrency data and analysis, including conditional formatting.

File Format
The Excel file will have the following columns:

Symbol: Cryptocurrency symbol (e.g., BTC)
Price (USD): The current price in USD
Market Cap: The market capitalization of the cryptocurrency
24h Volume: The 24-hour trading volume
24h Change (%): The 24-hour percentage change in price
The script will also update the top 5 cryptocurrencies by market cap and apply conditional coloring based on the 24-hour change percentage.

License
This project is licensed under the MIT License - see the LICENSE file for details.
