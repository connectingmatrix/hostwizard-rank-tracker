import datetime
import sys
import time
from django_cron import CronJobBase, Schedule
from .googlesheet.filter_data import FilterData
from .googlesheet.insert_data import DataInsertion
import pandas as pd
from multiprocessing import Pool
from joblib import Parallel, delayed

global p

# @background
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


def f(x):
    return x*x

def caller():
    print('Starting Multi process')
    with Pool(1) as pool:
        print(pool.map(f, [1,2], chunksize=1))


class MyCronJob(CronJobBase):
    RUN_EVERY_MINS = 1440  # Run every day

    schedule = Schedule(run_every_mins=RUN_EVERY_MINS)
    code = 'airbnb.MyCronJob'  # A unique code for your cron job

    def do(self):

        df_urls = pd.read_excel('urls.xlsx')
        print('My cronJob')
        p.map(f, range(len(df_urls)))
        print('finsihed loops')        
        sys.stdout.flush()



#if __name__ == '__main__':
cron = MyCronJob()

while True:
    p = Pool(2)
    start_time = datetime.datetime.now()
    print("Cron job started at:", start_time)
    cron.do()
    end_time = datetime.datetime.now()
    print("Cron job finished at:", end_time)
    time.sleep(cron.RUN_EVERY_MINS * 60)