import pandas as pd
from selenium import webdriver
from datetime import datetime, timedelta
from urllib.parse import urlparse, parse_qs
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
import time


class FilterData:
    def __int__(self):
        pass

    def find_available_dates(self, url):
        print("find_available_dates")
        #Set up the Chrome webdriver in headless mode
        options = webdriver.ChromeOptions()
        print("find_available_dates_1")
        #options.binary_location = 'googlechrome'
        options.add_argument('--disable-extensions')
        options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        # options.add_argument('--remote-debugging-port=9515')
        options.add_argument('--disable-setuid-sandbox')
        print("find_available_dates_2")
        
        driver = webdriver.Chrome(options=options)
        print("find_available_dates_3")
        
        wait = WebDriverWait(driver, 2000)
        print("find_available_dates_4")

        driver.get(url)
        print("find_available_dates_5")

        print("Driver Runing")
       
        rows_list = []
        check = 0
        while len(rows_list) < 360:
            
            try:
                wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR,
                                                        "div[role=application] > div:nth-child(2) > div > "
                                                        "div:nth-child(2) > div > table > tbody > tr > td["
                                                        "role=button]")))
            
                wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR,
                                                        'div[role=application] div:nth-child(2) > button['
                                                        'aria-label="Move forward to switch to the next month."]')))
                dates_rows = driver.find_elements(By.CSS_SELECTOR,
                                                "div[role=application] > div:nth-child(2) > div > div:nth-child(2) > "
                                                "div > table > tbody > tr > td > div[data-is-day-blocked=false]")
                
                # extract the dates and add them to the list
                #print(f"dates:{dates_rows}")
                if dates_rows:
                    for row in dates_rows:
                        date = row.get_attribute("data-testid")
                        date = date.split("-")[-1]
                        rows_list.append(date)

                    # click the next button to move to the next month wait.until(EC.invisibility_of_element_located((
                    # By.CSS_SELECTOR, 'div[data-testid="calendar-day-blocked"]')))
                    driver.find_element(By.CSS_SELECTOR,
                                        'div[role=application] div:nth-child(2) > button[aria-label="Move forward to switch '
                                        'to the next month."]').click()
                    time.sleep(2)
                    rows_list= [i for n, i in enumerate(rows_list) if i not in rows_list[:n]]
                else:
                    check += 1
                    if check >= 7:
                        break
                    driver.find_element(By.CSS_SELECTOR,'div[role=application] div:nth-child(2) > button[aria-label="Move forward to switch to the next month."]').click()
                    print(f"dates not found on first page now clicking next")
                    continue

            except Exception as ex:
        
                print("Exception occurs retrying in 10 seconds! Please Wait... ")
                # Refresh the page and wait for 10 seconds before trying again
                print('Exception Message: ',ex)
                driver.refresh()
                time.sleep(10)
                continue
        
        driver.quit()
        # convert list of dates to dataframe
        data = {'Available Dates': rows_list}
        df_available_dates = pd.DataFrame(data)
        df_available_dates['Available Dates'] = pd.to_datetime(df_available_dates['Available Dates'])
        df_available_dates
        # get checkin date from URL
        # parsed_url = urlparse(url)
        # query_params = parse_qs(parsed_url.query)
        # checkin_date = query_params['check_in'][0]
        
        checkin_date = str((df_available_dates['Available Dates'].iloc[0]).date())
        start_date = datetime.strptime(checkin_date, '%Y-%m-%d')
        end_date = start_date + timedelta(days=360)
        
        df_total_180_dates = df_available_dates[(df_available_dates['Available Dates'] >= start_date) & (df_available_dates['Available Dates'] < end_date)]
        
        end_date_40 = start_date + timedelta(days=40)
        # filter dates from checkin date to 180 days later
        df_40_dates = df_available_dates[(df_available_dates['Available Dates'] >= start_date) & (df_available_dates['Available Dates'] <= end_date_40)]
        
#       # Geting consecutive availabe dates in first 40 days
        df_40_consec_dates = self.find_consecutive_dates(df_40_dates)
        
        #checking for breakdown of the following: The weekend of Th - Su and weekday of M - Th
        df_40_split_dates = self.split_stays(df_40_consec_dates)
        
        
        # Finding dates geater than frist 40 dates

        df_40_greater = df_total_180_dates[df_total_180_dates['Available Dates'] > end_date_40]
 
 
        df_40_greater_consec_dates = self.find_consecutive_dates(df_40_greater)
        df_40_greater_split_dates = self.split_weekly(df_40_greater_consec_dates)
        
        df_final = pd.concat([df_40_split_dates, df_40_greater_split_dates], ignore_index=True)
        
        df_final = df_final[df_final['Check-In'] != df_final['Check-Out']]
        df_final['Check-In'] = pd.to_datetime(df_final['Check-In'])
        df_final['Check-Out'] = pd.to_datetime(df_final['Check-Out'])
        df_final.to_excel('input_filter.xlsx', index=False)
        print(f"input_filter xlsx file successfully created")
        
        return df_final



    def split_stays(self, df):
        stays = []
        for index, row in df.iterrows():
            checkin = datetime.strptime(str(row['Check-In']), '%Y-%m-%d')
            checkout = datetime.strptime(str(row['Check-Out']), '%Y-%m-%d')
            diff = (checkout - checkin).days

            if diff >= 7:
                stopover_days = [d for d in [checkin + timedelta(days=x) for x in range(1, diff)] 
                                if d.strftime('%A') in ['Sunday', 'Saturday', 'Thursday', 'Wednesday']]

                if len(stopover_days) == 0:
                    stays.append((checkin.strftime('%Y-%m-%d'), checkout.strftime('%Y-%m-%d')))
                else:
                    stopover_days.append(checkout)
                    stopover_days.insert(0, checkin)
                    for i in range(len(stopover_days) - 1):
                        stays.append((stopover_days[i].strftime('%Y-%m-%d'), stopover_days[i+1].strftime('%Y-%m-%d')))
            else:
                stays.append((checkin.strftime('%Y-%m-%d'), checkout.strftime('%Y-%m-%d')))
        df = pd.DataFrame(stays, columns=['Check-In', 'Check-Out'])

        # Convert dates to datetime format
        df['Check-In'] = pd.to_datetime(df['Check-In'])
        df['Check-Out'] = pd.to_datetime(df['Check-Out'])

        i = 0
        while i < len(df)-2:
            if df.loc[i+1, 'Check-In'] == df.loc[i, 'Check-Out'] and df.loc[i+1, 'Check-Out'].weekday() != 5:
                df.drop(index=i+1, inplace=True)
            df.reset_index(drop=True, inplace=True)
            i += 1
        df['Check-In'] = df['Check-In'].dt.strftime('%Y-%m-%d')
        df['Check-Out'] = df['Check-Out'].dt.strftime('%Y-%m-%d')
        return df
    
    def split_weekly(self,df):
        stays = []
        for index, row in df.iterrows():
            checkin = datetime.strptime(str(row['Check-In']), '%Y-%m-%d')
            checkout = datetime.strptime(str(row['Check-Out']), '%Y-%m-%d')
            diff = (checkout - checkin).days

            if diff >= 7:
                stopover_days = [d for d in [checkin + timedelta(days=x) for x in range(1, diff)] 
                                if d.strftime('%A') in ['Monday','Sunday']]

                if len(stopover_days) == 0:
                    stays.append((checkin.strftime('%Y-%m-%d'), checkout.strftime('%Y-%m-%d')))
                else:
                    stopover_days.append(checkout)
                    stopover_days.insert(0, checkin)
                    for i in range(len(stopover_days) - 1):
                        stays.append((stopover_days[i].strftime('%Y-%m-%d'), stopover_days[i+1].strftime('%Y-%m-%d')))
            else:
                stays.append((checkin.strftime('%Y-%m-%d'), checkout.strftime('%Y-%m-%d')))
        df = pd.DataFrame(stays, columns=['Check-In', 'Check-Out'])

        # Convert dates to datetime format
        df['Check-In'] = pd.to_datetime(df['Check-In'])
        df['Check-Out'] = pd.to_datetime(df['Check-Out'])

        i = 0
        while i < len(df)-2:
            if df.loc[i+1, 'Check-In'] == df.loc[i, 'Check-Out'] and df.loc[i+1, 'Check-Out'].weekday() != 5:
                df.drop(index=i+1, inplace=True)
            df.reset_index(drop=True, inplace=True)
            i += 1
        df['Check-In'] = df['Check-In'].dt.strftime('%Y-%m-%d')
        df['Check-Out'] = df['Check-Out'].dt.strftime('%Y-%m-%d')
        return df


    def find_consecutive_dates(self, df):
        """
        Find the consecutive dates and create an input filter file for the Airbnb search based on the available dates DataFrame.

        Parameters:
        df (pandas.DataFrame): DataFrame containing the available dates.

        Returns:
        pandas.DataFrame: DataFrame containing the input filters.
        """
        # Create empty lists to store start and end dates of consecutive dates
        start_dates = []
        end_dates = []

        # Initialize start and end dates to the first date in the column
        start_date = df['Available Dates'].iloc[0]
        end_date = df['Available Dates'].iloc[0]

        # Loop through dataframe and find consecutive dates
        for i in range(1, len(df)):
            if (df['Available Dates'].iloc[i] - end_date).days == 1:
                end_date = df['Available Dates'].iloc[i]
            else:
                start_dates.append(start_date)
                end_dates.append(end_date)
                start_date = df['Available Dates'].iloc[i]
                end_date = df['Available Dates'].iloc[i]

        # Append last consecutive date range to start_dates and end_dates lists
        start_dates.append(start_date)
        end_dates.append(end_date)
    #
        date_ranges = pd.DataFrame({'Check-In': start_dates, 'Check-Out': end_dates})
        
        # Optional: Convert the start_date and end_date columns to strings
        date_ranges['Check-In'] = date_ranges['Check-In'].dt.strftime('%Y-%m-%d')
        date_ranges['Check-Out'] = date_ranges['Check-Out'].dt.strftime('%Y-%m-%d')

        return date_ranges


    
# if __name__ == "__main__":
#     filter_instance = FilterData()
#     filter_instance.find_available_dates()
