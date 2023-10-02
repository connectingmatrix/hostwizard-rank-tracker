import schedule
import time
import subprocess
import datetime

# Define the function to run the job
def run_job():
    # now = datetime.datetime.now()
    # print(f"Running job at {now}")

    # Replace 'your_script.py' with the name of the script you want to run
    subprocess.Popen(["python", "scrapper.py", "0"])

# Schedule the job to run every day at midnight
run_job()
schedule.every().day.at("00:00").do(run_job)

# Infinite loop to keep the script running
while True:
    schedule.run_pending()
    time.sleep(1)
