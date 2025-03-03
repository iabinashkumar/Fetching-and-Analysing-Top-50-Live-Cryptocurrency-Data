import requests
import pandas as pd
import xlwings as xw
import time
from datetime import datetime

def fetch_cryptocurrency_data():
    url = 'https://api.coingecko.com/api/v3/coins/markets'
    params = {
        'vs_currency': 'usd',
        'order': 'market_cap_desc',
        'per_page': 50,
        'page': 1,
        'sparkline': False
    }
    response = requests.get(url, params=params)
    data = response.json()
    return data

def create_dataframe(crypto_data):
    dataframe = pd.DataFrame(crypto_data)
    print("DataFrame columns:", dataframe.columns)  # Debugging line to display columns

    dataframe = dataframe[['name', 'symbol', 'current_price', 'market_cap', 'total_volume', 'price_change_percentage_24h']]
    dataframe.columns = ['Name', 'Symbol', 'Current Price (USD)', 'Market Cap', 'Trading Volume', 'Price Change (24h, %)']
    dataframe.index = dataframe.index + 1  # Start indexing from 1
    return dataframe

def analyze_crypto_data(dataframe):
    top_five_cryptos = dataframe.nlargest(5, 'Market Cap')
    top_five_cryptos.index = range(1, 6)  # Ensure top crypto numbering starts from 1
    avg_price = dataframe['Current Price (USD)'].mean()
    highest_price_change = dataframe['Price Change (24h, %)'].max()
    lowest_price_change = dataframe['Price Change (24h, %)'].min()

    # Display the results on the screen
    print("Top 5 cryptocurrencies by market cap:")
    print(top_five_cryptos)
    print("\nAverage price of top 50 cryptocurrencies: ${:.2f}".format(avg_price))
    print("\nHighest price change in 24 hours: {:.2f}%".format(highest_price_change))
    print("Lowest price change in 24 hours: {:.2f}%".format(lowest_price_change))

def update_excel(dataframe, file_path):
    # Create a new workbook if it doesn't exist, otherwise open the existing workbook
    try:
        workbook = xw.Book(file_path)
    except FileNotFoundError:
        workbook = xw.Book()
        workbook.save(file_path)
    
    sheet = workbook.sheets['Sheet1']  # Select the first sheet
    sheet.clear_contents()  # Clear the sheet contents
    sheet.range('A1').value = dataframe  # Write the DataFrame to the sheet

    # Show the update time in the Excel sheet
    update_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    sheet.range('G1').value = f"Last updated: {update_time}"
    workbook.save(file_path)  # Save the workbook

def main():
    file_path = 'live_crypto_data.xlsx'
    while True:
        crypto_data = fetch_cryptocurrency_data()
        dataframe = create_dataframe(crypto_data)
        analyze_crypto_data(dataframe)
        update_excel(dataframe, file_path)
        print(f"Excel sheet updated at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        time.sleep(300)  # Update every 5 minutes

if __name__ == "__main__":
    main()
