import requests
import time
import openpyxl
import os
import logging
import win32com.client
import psutil

API_URL = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest"
API_KEY = "5992c351-18aa-492a-8d03-cb8c6c27dc46" 
FETCH_INTERVAL = 10 
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(message)s')

HEADERS = {
    "Accepts": "application/json",
    "X-CMC_PRO_API_KEY": API_KEY,
}

def fetch_live_data():
    logging.info("Fetching live cryptocurrency data...")
    try:
        response = requests.get(API_URL, headers=HEADERS, params={"limit": 50})
        response.raise_for_status()
        data = response.json()

        if "data" not in data:
            logging.error("No 'data' field found in the API response.")
            return []
        
        return data["data"]
    except requests.exceptions.RequestException as e:
        logging.error(f"Error fetching data from API: {e}")
        return []

def process_data(data):
    processed = []
    for item in data:
        processed.append({
            "Name": item["name"],
            "Symbol": item["symbol"],
            "Price (USD)": item["quote"]["USD"]["price"],
            "Market Cap": item["quote"]["USD"]["market_cap"],
            "24h Volume": item["quote"]["USD"]["volume_24h"],
            "24h Change (%)": item["quote"]["USD"]["percent_change_24h"]
        })
    return processed

def analyze_data(data):
    top_5 = sorted(data, key=lambda x: x["Market Cap"])[:5]
    average_price = sum([item["Price (USD)"] for item in data]) / len(data)
    highest_change = max(data, key=lambda x: x["24h Change (%)"])
    lowest_change = min(data, key=lambda x: x["24h Change (%)"])

    analysis = {
        "Top 5 Cryptocurrencies": top_5,
        "Average Price of Top 50": average_price,
        "Highest 24h Change": highest_change,
        "Lowest 24h Change": lowest_change
    }

    return analysis

def is_excel_file_open(excel_app, file_name):
    try:
        workbooks = excel_app.Workbooks
        for wb in workbooks:
            if os.path.abspath(file_name) == wb.FullName:
                return wb
        return None
    except Exception as e:
        logging.error(f"Error checking open workbooks: {e}")
        return None

def open_excel_file(file_name):
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        workbook = is_excel_file_open(excel, file_name)
        if workbook is None:
            workbook = excel.Workbooks.Open(os.path.abspath(file_name))
        excel.Visible = True
        return workbook, excel
    except Exception as e:
        logging.error(f"Error opening Excel file: {e}")
        return None, None

def update_excel(file_name, data):
    logging.info("Updating the Excel sheet with new data...")

    try:
        workbook, excel = open_excel_file(file_name)
        if workbook:
            sheet = workbook.Sheets(1)
            sheet.Name = "Cryptocurrency Data"

            sheet.Rows("2:1048576").Delete()

            for i, item in enumerate(data, start=2):
                sheet.Cells(i, 1).Value = item["Name"]
                sheet.Cells(i, 2).Value = item["Symbol"]
                sheet.Cells(i, 3).Value = round(item["Price (USD)"], 2)
                sheet.Cells(i, 4).Value = round(item["Market Cap"], 2)
                sheet.Cells(i, 5).Value = round(item["24h Volume"], 2)
                sheet.Cells(i, 6).Value = round(item["24h Change (%)"], 2)

                change_cell = sheet.Cells(i, 6)
                if item["24h Change (%)"] > 0:
                    change_cell.Interior.Color = 0x90EE90  
                elif item["24h Change (%)"] < 0:
                    change_cell.Interior.Color = 0xFFB6C1  

            sheet.Cells(len(data) + 2, 1).Value = "Market Cap"
            sheet.Cells(len(data) + 2, 1).Font.Bold = True
            sheet.Cells(len(data) + 2, 1).Font.Size = 12

            top_5 = sorted(data, key=lambda x: x["Market Cap"])[:5]
            for i, item in enumerate(top_5, start=len(data) + 3):
                sheet.Cells(i, 1).Value = item["Name"]
                sheet.Cells(i, 2).Value = item["Symbol"]
                sheet.Cells(i, 3).Value = round(item["Price (USD)"], 2)
                sheet.Cells(i, 4).Value = round(item["Market Cap"], 2)
                sheet.Cells(i, 5).Value = round(item["24h Volume"], 2)
                sheet.Cells(i, 6).Value = round(item["24h Change (%)"], 2)

                change_cell = sheet.Cells(i, 6)
                if item["24h Change (%)"] > 0:
                    change_cell.Interior.Color = 0x90EE90
                elif item["24h Change (%)"] < 0:
                    change_cell.Interior.Color = 0xFF9999

                market_cap_cell = sheet.Cells(i, 4)
                if i == len(data) + 3:
                    market_cap_cell.Interior.Color = 0xADD8E6
                elif i == len(data) + 4:
                    market_cap_cell.Interior.Color = 0xFFFFE0
                elif i == len(data) + 5:
                    market_cap_cell.Interior.Color = 0xF08080
                elif i == len(data) + 6:
                    market_cap_cell.Interior.Color = 0xE0FFFF
                elif i == len(data) + 7:
                    market_cap_cell.Interior.Color = 0xFFB6C1

            workbook.Save()
            logging.info("Excel updated successfully!")

        else:
            logging.error(f"Failed to find or open the Excel workbook: {file_name}")
    except Exception as e:
        logging.error(f"Error updating Excel: {e}")


def main():
    excel_file_name = input("Enter the Excel file name (e.g., k1.xlsx): ").strip()

    if not os.path.isfile(excel_file_name):
        logging.error(f"The file '{excel_file_name}' does not exist. Please provide a valid file.")
        return

    logging.info("Starting live cryptocurrency data fetch...")

    while True:
        data = fetch_live_data()
        if data:
            processed_data = process_data(data)
            analysis = analyze_data(processed_data)
            logging.info(f"Analysis Results: {analysis}")
            update_excel(excel_file_name, processed_data)
        else:
            logging.warning("No data fetched. Retrying in the next interval.")
        
        time.sleep(FETCH_INTERVAL)

if __name__ == "__main__":
    main()
