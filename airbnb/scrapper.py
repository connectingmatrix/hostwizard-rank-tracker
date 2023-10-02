from multiprocessing import Pool
import sys
from googlesheet.filter_data import FilterData
from googlesheet.insert_data import DataInsertion
import pandas as pd
from multiprocessing import Pool
import re

def get_airbnb_id_from_url(url):
    # The regular expression looks for '/rooms/' followed by one or more digits
    match = re.search(r'/rooms/(\d+)', url)
    if match:
        return int(match.group(1))
    else:
        return None

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

    airbnb_id = get_airbnb_id_from_url(url)

    print('Calling main')
    insert_data.main(spreadsheet_name, hotel_name, description, location, price_log_sheet, airbnb_id)


if __name__ == '__main__':
    
    filterAndInsert(int(sys.argv[1]))
    # with Pool(2) as p:
    #     p.map(filterAndInsert, range(len(df_urls)))
    #     # print(p.map(filterAndInsert, range(len(df_urls))))