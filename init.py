from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
import time
import csv
import numpy as np
import os
import pandas as pd
import matplotlib.pyplot as plt


# Set the download directory
download_dir = 'C:/Users/hoorm/Downloads'
# Get a list of all files in the download directory before the download
before = os.listdir(download_dir)

# Specify the path to the ChromeDriver executable
chromedriver_path = 'chromedriver.exe'  # Replace with the actual path

# Specify the path to your Chrome profile
chrome_data_path = "C:/Users/hoorm/AppData/Local/Google/Chrome/User Data/"
chrome_profile_path = "Profile 4" 


# Specify the URL you want to navigate to
url = 'https://business.facebook.com/latest/insights/results?business_id=693823586008930&asset_id=188285954362840&ad_account_id=120200701245850619&entity_type=FB_PAGE'

# Create Chrome options and set the profile path
chrome_options = Options()
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument(f'--user-data-dir={chrome_data_path}')
chrome_options.add_argument(f'--profile-directory={chrome_profile_path}')
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])

# Create a new instance of the Chrome driver with the specified options
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# Navigate to the specified URL
driver.get(url)
time.sleep(8)

# Wait for the element to be clickable
matches = pyautogui.locateAllOnScreen('img/Export.png', confidence=0.8)
for match in matches:
    print(match)
    pyautogui.click(match)
    time.sleep(1)
    download = pyautogui.locateOnScreen('img/csv.png', confidence=0.95)
    pyautogui.click(download)
    time.sleep(1)

# Close the browser
driver.quit()

# Get a list of all files in the download directory after the download
after = os.listdir(download_dir)

# Get a list of the downloaded files
downloaded_files = [f for f in after if f not in before]

# # Read the contents of the downloaded files and then delete them
# for filename in downloaded_files:
#     if filename.endswith(".csv"):
#         with open(os.path.join(download_dir, filename), newline='') as csvfile:
#             reader = csv.reader(csvfile)
#             for row in reader:
#                 print(', '.join(row))
#         os.remove(os.path.join(download_dir, filename))  # delete the file

# for filename in downloaded_files:
#     if filename.endswith(".csv"):
#         # Read the CSV file into a pandas DataFrame
#         df = pd.read_csv(os.path.join(download_dir, filename), sep=',', encoding='utf-16')

#         # Split the DataFrame into separate DataFrames for each table
#         tables = []
#         start = 0
#         for i in range(1, len(df)):
#             if pd.isna(df.iloc[i, 0]):
#                 tables.append(df.iloc[start:i])
#                 start = i + 1
#         tables.append(df.iloc[start:])

#         # Draw a graph for each table
#         for i, table in enumerate(tables):
#             table = table.dropna()  # remove rows with missing values
#             table.columns = table.iloc[0]  # set the column names

#             # Check if the first row can be converted to numeric values
#             first_row_can_be_numeric = not np.isnan(pd.to_numeric(table.iloc[1, 1], errors='coerce'))

#             # If the first row cannot be converted to numeric values, drop it
#             if not first_row_can_be_numeric:
#                 table = table.iloc[1:]

#             table[table.columns[1]] = pd.to_numeric(table[table.columns[1]], errors='coerce')  # convert the second column to numeric values
#             table.plot(x=table.columns[0], y=table.columns[1], kind='line', title=table.columns[1])
#             plt.savefig(os.path.join(download_dir, f'{filename}_{i}.png'))  # save the graph as a PNG image

#         os.remove(os.path.join(download_dir, filename))  # delete the file