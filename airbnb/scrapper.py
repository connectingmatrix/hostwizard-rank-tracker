from multiprocessing import Pool
import sys
from googlesheet.filter_data import FilterData
from googlesheet.insert_data import DataInsertion
import pandas as pd
from multiprocessing import Pool

def filterAndInsert(x):

    df_urls = pd.read_excel('urls.xlsx')
    first_row = df_urls.iloc[x]
    spreadsheet_name = 'Ranking Data - Sunshine Lux Rentals'
    price_log_sheets = ['Price Change Log(Villa in Hollywood)', 'Price Change Log(West Palm Beach)']

    url = first_row[0]
    hotel_name = first_row[1]
    description = first_row[2]
    location = first_row[3]
    price_log_sheet = price_log_sheets[x]

    print('[POOL]',url, spreadsheet_name, hotel_name, description, location, price_log_sheet)        
    sys.stdout.flush()

    filter_data = FilterData()
    filter_data.find_available_dates(url)

    # Inserting data in Spreadsheets
    print("Inserting Data..", spreadsheet_name)
    sys.stdout.flush()
    insert_data = DataInsertion()
    insert_data.main(spreadsheet_name, hotel_name, description, location, price_log_sheet)



if __name__ == '__main__':
    df_urls = pd.read_excel('urls.xlsx')

    with Pool(2) as p:
        print(p.map(filterAndInsert, range(len(df_urls))))