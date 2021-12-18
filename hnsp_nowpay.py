import os
from os import path
import shutil
import json
import tkinter
from tkinter import filedialog
import pandas as pd
import json
from openpyxl.workbook import Workbook
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.action_chains import ActionChains
import time

root = tkinter.Tk()
root.withdraw()

# Open xlsx file
open_sheet = path.exists("cache/opened_sheet.json")
if open_sheet == True :
  opened_sheet_file_path = "cache/opened_sheet.json"
  json_file = open(opened_sheet_file_path)
  data = json.load(json_file)
  xlsx_sheet_check = path.exists(data ['xlsx_file_path'])
  if xlsx_sheet_check == True :
    xlsx_file_path = data ['xlsx_file_path']
  else :
    shutil.rmtree('cache', ignore_errors=True)
    xlsx_file_path = filedialog.askopenfilename(title="Open Excel-XLSX File")
    cache_path = os.path.join(str(os.getcwd()), "cache")
    dictionary = {"xlsx_file_path" : xlsx_file_path}
    json_object = json.dumps(dictionary, indent = 1)
    with open("cache/opened_sheet.json", "w") as outfile:
      outfile.write(json_object)
else :
  xlsx_file_path = filedialog.askopenfilename(title="Open Excel-XLSX File")
  cache_path = os.path.join(str(os.getcwd()), "cache")
  os.mkdir(cache_path)
  dictionary = {"xlsx_file_path" : xlsx_file_path}
  json_object = json.dumps(dictionary, indent = 1)
  with open("cache/opened_sheet.json", "w") as outfile:
    outfile.write(json_object)

# read imported xlsx file path using pandas
input_workbook = pd.read_excel(xlsx_file_path, sheet_name = 'Sheet1', usecols = 'E:I', dtype=str)
input_workbook.head()

# read total number of rows present in xlsx
number_of_rows = len(input_workbook.index)

# Opening JSON file & returns JSON object as a dictionary
json_file = open('settings.json')
settings_data = json.load(json_file)

input_workbook_cc_number = input_workbook['Card Number'].values.tolist()
input_workbook_cvv_number = input_workbook['CVV'].values.tolist()
input_workbook_expiry_number = input_workbook['Expiry'].values.tolist()
input_workbook_atm_pin = input_workbook['ATM pin'].values.tolist()
input_workbook_desk_number = input_workbook['Desk'].values.tolist()

# get-output sheet to append output
output_sheet = path.exists("Output.xlsx")
if output_sheet == True :
  output_sheet_file_path = "Output.xlsx"
else :
  output_headers= ['FirstName','LastName', 'Mobile', 'Email','Amount', 'CardNumber', 'CVV', 'Expiry', 'ATM pin', 'No.of Transactions', 'Desk', "Desk Holder"]
  overall_output = Workbook()
  page = overall_output.active
  page.append(output_headers)
  overall_output.save(filename = 'Output.xlsx')
  output_sheet_file_path = "Output.xlsx"

def cal():
  global output_cc_number
  global done_transactions_wb_1
  global h
  output_load_wb_2 = pd.read_excel(output_sheet_file_path, sheet_name = 'Sheet', usecols = 'F', dtype=str)
  output_load_wb_2.head()
  output_cc_number = output_load_wb_2['CardNumber'].values.tolist()
  output_load_wb_1 = pd.read_excel(output_sheet_file_path, sheet_name = 'Sheet', usecols = 'J', dtype=int)
  output_load_wb_1.head()
  done_transactions_wb_1 = output_load_wb_1['No.of Transactions'].values.tolist()
  h = len(output_load_wb_1.index) - 1
  print (output_cc_number[h],done_transactions_wb_1[h])

def cc_expiry():
  global expiry_month
  global expiry_year
  global expiry_year1
  global expiry_year2
  global expiry_year3
  global expiry_year4
  workbook_expiry_month = input_workbook_expiry_number[x]
  workbook_expiry_year = input_workbook_expiry_number[x]
  expiry_month = workbook_expiry_month[:2]
  expiry_year = workbook_expiry_year[5:]
  expiry_year1 = workbook_expiry_year[3]
  expiry_year2 = workbook_expiry_year[4]
  expiry_year3 = workbook_expiry_year[5]
  expiry_year4 = workbook_expiry_year[6]

def textbox_field(xpath, timeout_time, send_keys_data):
  try :
    WebDriverWait(driver, timeout=timeout_time).until(ec.visibility_of_element_located((By.XPATH, xpath)))
  except TimeoutException:
    timeout_exception()
  else :
    textbox_elements = driver.find_element_by_xpath (xpath)
    textbox_elements.send_keys(send_keys_data)

def button_field(button_xpath, timeout_time):
  try :
    WebDriverWait(driver, timeout=timeout_time).until(ec.visibility_of_element_located((By.XPATH, button_xpath)))
  except TimeoutException:
    timeout_exception()
  else :
    page_button = driver.find_element_by_xpath (button_xpath)
    page_button.click()


def textbox_field_click(xpath):
  try :
    time.sleep(0.50)
    act.click(driver.find_element_by_xpath (xpath)).perform()
    #WebDriverWait(driver, timeout=timeout_time).until(ec.visibility_of_element_located((By.XPATH, xpath)))
  except NoSuchElementException:
    timeout_exception()
  else :
    act.click(driver.find_element_by_xpath (xpath)).perform()

def start_link():
  driver.get("https://hnsp.nowpay.co.in/")

def pageone():
    textbox_field('//*[@id="paymentmaster-first_name"]', 8, settings_data['first_name'])
    textbox_field('//*[@id="paymentmaster-last_name"]', 8, settings_data['last_name'])
    textbox_field('//*[@id="paymentmaster-email"]', 8, settings_data['email_id'])
    textbox_field('//*[@id="paymentmaster-phone"]', 8, settings_data['registered_mobile_no'])
    textbox_field('//*[@id="paymentmaster-address"]', 8, settings_data['address'])
    textbox_field('//*[@id="paymentmaster-city"]', 8, settings_data['address'])
    button_field('//*[@id="subm"]', 8)

def pagetwo():
    textbox_field_click('//*[@id="wrap"]/div[3]/div[2]/div/div[2]/div[4]/div[1]/div[5]/div[1]/div[1]/div[1]/div/input')
    textbox_field('//*[@id="wrap"]/div[3]/div[2]/div/div[2]/div[4]/div[1]/div[5]/div[1]/div[1]/div[1]/div/input', 8, input_workbook_cc_number[x])
    
    textbox_field_click('//*[@id="wrap"]/div[3]/div[2]/div/div[2]/div[4]/div[1]/div[5]/div[1]/div[1]/div[3]/div/div[1]/div/div/input')
    textbox_field('//*[@id="wrap"]/div[3]/div[2]/div/div[2]/div[4]/div[1]/div[5]/div[1]/div[1]/div[3]/div/div[1]/div/div/input', 8, expiry_month + expiry_year)

    textbox_field_click('//*[@id="wrap"]/div[3]/div[2]/div/div[2]/div[4]/div[1]/div[5]/div[1]/div[1]/div[3]/div/div[2]/div/div/input')
    textbox_field('//*[@id="wrap"]/div[3]/div[2]/div/div[2]/div[4]/div[1]/div[5]/div[1]/div[1]/div[3]/div/div[2]/div/div/input', 8, input_workbook_cvv_number[x])
    button_field('//*[@id="wrap"]/div[3]/div[2]/div/div[2]/div[4]/div[1]/div[5]/div[1]/div[3]/input', 8)

def pagethree():
    button_field('//*[@id="tab-B-label"]/span', 8)
    textbox_field('//*[@id="expDate"]', 8, expiry_month)
    textbox_field('//*[@id="expDate"]', 8, expiry_year1)
    textbox_field('//*[@id="expDate"]', 8, expiry_year2)
    textbox_field('//*[@id="expDate"]', 8, expiry_year3)
    textbox_field('//*[@id="expDate"]', 8, expiry_year4)
    textbox_field('//*[@id="pin"]', 8, input_workbook_atm_pin[x])
    button_field('//*[@id="submitButtonIdForPin"]', 8)
    time.sleep(1000)


def output_save():
  entry_list = [[settings_data['first_name'], settings_data['last_name'], settings_data['registered_mobile_no'], settings_data['email_id'], settings_data['payable_amount'], input_workbook_cc_number[x], input_workbook_atm_pin[x], input_workbook_cvv_number[x], input_workbook_expiry_number[x], z+1, int(input_workbook_desk_number[x]), settings_data["desk_holder"]]]
  output_wb = load_workbook(output_sheet_file_path)
  page = output_wb.active
  for info in entry_list:
      page.append(info)
  output_wb.save(filename='Output.xlsx')

def timeout_exception():
    start_link()
    pageone()
    cc_expiry()
    time.sleep(2)
    pagetwo()
    pagethree()
    print ("exception")

def whole_work():
    start_link()
    pageone()
    cc_expiry()
    time.sleep(2)
    pagetwo()
    pagethree()


caps = DesiredCapabilities().CHROME
caps["pageLoadStrategy"] = "none"
#caps["pageLoadStrategy"] = "eager"
#caps["pageLoadStrategy"] = "normal"
driver=webdriver.Chrome(desired_capabilities=caps, executable_path="chromedriver.exe")
driver.maximize_window()
act = ActionChains(driver)
try:
  cal()
except IndexError:
  for x in range (0 , number_of_rows):
    for z in range (0, int(settings_data['number_of_time_transactions_per_card'])):
      whole_work()
else:
  last_txncard =  input_workbook[input_workbook['Card Number'] == output_cc_number[h]].index[0]
  for x in range (last_txncard , number_of_rows):
    for z in range (done_transactions_wb_1[h], int(settings_data['number_of_time_transactions_per_card'])):
      whole_work()
    done_transactions_wb_1[h] = 0

driver.quit()
