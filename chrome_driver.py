import time
import sqlite3
import pandas as pd
import os

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from sqlite3 import Error

download_dir = os.getcwd()
xls_file = download_dir+"\skill_test_data.xlsx"
table_name = 'pivot_table'
print(download_dir)

def connect_webdriver():
  options = webdriver.ChromeOptions()
  prefs = {"download.default_directory" : download_dir}
  options.add_experimental_option("prefs",prefs)

  # options.add_argument("start-maximized")
  options.add_experimental_option('excludeSwitches', ['enable-logging']) # disable device_event_log
  options.binary_location = r"C:\Program Files\Google\Chrome Beta\Application\chrome.exe" # version 111
  chrome_driver_binary = r"D:/chromedriver_win32/chromedriver.exe" # beta version
  driver = webdriver.Chrome(chrome_driver_binary, chrome_options=options)

   # Part 1
  driver.get('https://jobs.homesteadstudio.co/data-engineer/assessment/download')
  downloadcsv = driver.find_element(By.CSS_SELECTOR,'.wp-block-button')
  downloadcsv.click()
  time.sleep(5)
  driver.close()


def generate_pivot_table():
  conn = None
  try:
    xls = pd.ExcelFile(xls_file)
    df = pd.read_excel(xls, "data")
    # cols_to_display = ["Spend", "Attributed Rev (1d)", "Visits", "New Visits", "Transactions (1d)", "Signups (1d)"]
    pivot_table = df.groupby(["Platform (Northbeam)"]).sum(numeric_only=True)
    pt = pivot_table[["Spend", "Attributed Rev (1d)", "Visits", "New Visits", "Transactions (1d)", "Email Signups (1d)"]]
    sorted_data = pt.sort_values(by=['Attributed Rev (1d)','Transactions (1d)'], ascending=False)
    sorted_data.loc["Grand Total"] = sorted_data.sum()
    # print(sorted_data)
    # print(sorted_data.info())

    conn = sqlite3.connect('mydb.sqlite')
    query = f'Create table if not Exists {table_name} (row_labels text, spend_sum real, attributed_rev_1d_sum real, imprs_sum real, visits_sum real, \
              new_visits_sum real, transactions_1d_sum real, email_signups_1d_sum real)'
    conn.execute(query)
    sorted_data.to_sql(table_name,conn,if_exists='replace',index=True)
    conn.commit()

    select = pd.read_sql('select * from pivot_table',conn)
    print(select)

  except Exception as e:
    print(e)

  finally:
    if conn:
      conn.close()

  
if __name__ == '__main__':
  try:
    connect_webdriver()
    generate_pivot_table()

  except Error as e:
    print(e)