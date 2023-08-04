import datetime
import time
from django_cron import CronJobBase, Schedule
from .googlesheet.filter_data import FilterData
from .googlesheet.insert_data import DataInsertion
import pandas as pd

class MyCronJob(CronJobBase):
    RUN_EVERY_MINS = 1440  # Run every day

    schedule = Schedule(run_every_mins=RUN_EVERY_MINS)
    code = 'airbnb.MyCronJob'  # A unique code for your cron job

    def do(self):
        # Add your code logic here
        
        #  Creating input file filter file
        spreadsheet_name = 'Ranking Data - Sunshine Lux Rentals'
        price_log_sheets = ['Price Change Log(Villa in Hollywood)', 'Price Change Log(West Palm Beach)']
        df_urls = pd.read_excel('urls.xlsx')
    
        for a in range(len(df_urls)): 
        
            first_row = df_urls.iloc[a]
            url = first_row[0]
            hotel_name = first_row[1]
            description = first_row[2]
            location = first_row[3]
            price_log_sheet = price_log_sheets[a]

            filter_data = FilterData()
            filter_data.find_available_dates(url)

            # Inserting data in Spreadsheets
            print("Inserting Data..")
            insert_data = DataInsertion()
            insert_data.main(spreadsheet_name, hotel_name, description, location, price_log_sheet)


cron = MyCronJob()
while True:
    start_time = datetime.datetime.now()
    print("Cron job started at:", start_time)
    cron.do()
    end_time = datetime.datetime.now()
    print("Cron job finished at:", end_time)
    time.sleep(cron.RUN_EVERY_MINS * 60)
