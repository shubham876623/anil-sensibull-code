import requests
import pandas as pd
from io import BytesIO
# Anil sir sensibull credential
# Id BV0780
# Pass KrishnaNo.1
from selenium import webdriver
from selenium.webdriver.chrome.options import  Options
from random import seed
from random import random
from time import process_time, sleep, time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common import keys
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.common.keys import Keys
import pandas as pd
from bs4 import BeautifulSoup
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import urllib.request
import csv
from urllib.request import urlopen
from selenium.webdriver.support.select import Select
import re
from selenium.webdriver.common.alert import Alert


import gspread
from oauth2client.service_account import ServiceAccountCredentials
chrome_options = webdriver.ChromeOptions()


scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

credentials = ServiceAccountCredentials.from_json_keyfile_name('sensibull-bot-d9c2b6cc9a70.json', scope)
client = gspread.authorize(credentials)

spreadsheet = client.open('Sensibull Bot')
gc=gspread.service_account(filename='sensibull-bot-d9c2b6cc9a70.json')
sh=gc.open_by_url('https://docs.google.com/spreadsheets/d/1qe0uq06rYaOqyH0Q82FkYrCrp37oYGv2snXc2QHdOBU/edit')
# worksheet_list = sh.worksheets()

worksheet=sh.get_worksheet(0)
data=[]
prefs = {"profile.default_content_setting_values.notifications" : 2}
driver=webdriver.Chrome(ChromeDriverManager().install(),chrome_options=chrome_options)
driver.get('https://web.sensibull.com/login')
driver.maximize_window()
time.sleep(2)
driver.find_element_by_xpath("//*[contains(text(),'Zerodha')]").click()
time.sleep(4)

driver.find_element_by_id('userid').send_keys('')
driver.find_element_by_id('password').send_keys('')

driver.find_element_by_xpath('//button[@type="submit"]').click()

otp=input(str("Enter The Otp please : "))
driver.find_element_by_xpath("//input[@id='totp']").send_keys(otp)
time.sleep(1)
driver.find_element_by_xpath('//button[@type="submit"]').click()
time.sleep(2)

driver.get('https://web.sensibull.com/positions')

time.sleep(10)
while True :
    
    header = ["Scrip Code","Month","Type","Strike Price",'LTP' ,'Stock prices' ,"Last Updated"]
    with open('test_output.csv', 'w', encoding='UTF8') as f:
        writer = csv.writer(f)

    # write the header
        writer.writerow(header)
    soup=BeautifulSoup(driver.page_source,'html.parser')
    
    all_div=soup.find_all('div',{'class':'MuiCollapse-wrapperInner'})
    for one_div in all_div:
        try: 
            scrip_code=one_div.find('div',{'class':'style__InstrumentTickerStatsWrapper-sc-p1an74-6 hdHcuF instrument-ticker'})
    
            
            all_rows_in_one_div=one_div.find('tbody',{'class':'MuiTableBody-root'}).find_all('tr')
            for single_row in all_rows_in_one_div :
                month=single_row.find('span',{'class':'month'})
                
                Type=single_row.find('span',{'class':'symbol'})
                ltp=single_row.find('td',{'accessor':'ltp'})
                
            
                if len(re.findall('\d+',Type.text.split('th')[1] ))<1:
                    strike_prices='.'
                else:
                    strike_prices=re.findall('\d+',Type.text.split('th')[1])[0]
               
                data.append(scrip_code.text.split()[0])
                data.append(month.text)
                data.append(Type.text.split()[1])
                data.append(strike_prices)
                data.append(ltp.text)
                data.append(scrip_code.text.split()[1])
                t = time.localtime()
                current_time = time.strftime("%H:%M:%S", t)
                data.append(current_time)
                  
                df=pd.DataFrame([data])
                df.to_csv("test_output2.csv",mode="a",header=False,index=False)
                
                data.clear()
        except:
            pass
    spreadsheet = client.open('Sensibull Bot')

    with open('test_output.csv', 'r') as file_obj:
        content = file_obj.read()
        client.import_csv(spreadsheet.id, data=content)  
    time.sleep(300)    
    worksheet.clear()
