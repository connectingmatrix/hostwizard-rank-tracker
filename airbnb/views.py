import time
from django.shortcuts import render
from django.http import HttpResponse, JsonResponse
from .models import GoogleSheet
from .googlesheet.filter_data import FilterData
from .googlesheet.insert_data import DataInsertion
# from googlesheet import insert_data
import subprocess
from rest_framework.decorators import api_view
from rest_framework.response import Response
from threading import Thread
import pandas as pd

def homepage(request):
    return render(request, 'airbnb/home.html')


# Create your views here.
def run_api(request):
    
    print("run_api")
    
    #  Creating input file filter file
    
    spreadsheet_name='Ranking Data - Sunshine Lux Rentals'
    price_log_sheets = ['Price Change Log(Villa in Hollywood)', 'Price Change Log(West Palm Beach)']
    df_urls = pd.read_excel('urls.xlsx')
    
    for a in range(len(df_urls)): 
        
        first_row = df_urls.iloc[a]
        url= first_row[0]
        
        hotel_name= first_row[1]
        description= first_row[2]
        location = first_row[3]
        price_log_sheet= price_log_sheets[a]


        filter_data = FilterData()
        filter_data.find_available_dates(url)
        
        
        # Inserting data in Spreadsheets
        print("Inserting Data..")

        insert_data = DataInsertion()
        insert_data.main(spreadsheet_name,hotel_name, description,location,price_log_sheet)

        # Inserting values into price log

    time.sleep(5)
    # return a response to the client
    return HttpResponse('https://docs.google.com/spreadsheets/d/1MEDjSL9AenhQ_R7rYPSbr1G5jpikbuuupfHjCnGD9LM/edit#gid=1133271900')

    

def run_bot(request):
    if request.method == 'GET':
        thread = Thread(target = run_api(request),)
        thread.start()
    
    else:
        pass

