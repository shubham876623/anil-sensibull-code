import requests
import pandas as pd
from io import BytesIO
from selenium import webdriver
from selenium.webdriver.chrome.options import  Options
from random import seed
from random import random
from time import process_time, sleep
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
#from emailotp import get_otp
from selenium.webdriver.common.alert import Alert
from datetime import date
import pyautogui as pg
import time
import re
import win32con
import win32gui                  

#driver=webdriver.Chrome(ChromeDriverManager(version='114.0.5735.90').install())
from selenium.webdriver.chrome.service import Service
service = Service(r"chromedriver.exe")
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)


driver.get('https://stockexcel.com/Login.aspx')
driver.find_element(By.ID ,'ContentPlaceHolder1_txtEmailID').send_keys('documents@kariwala.com')
driver.find_element(By.ID ,'ContentPlaceHolder1_txtPassword').send_keys('Sumit@3363')

driver.find_element(By.ID ,'ContentPlaceHolder1_btnLogin').click()
time.sleep(5)
email_input=driver.find_element(By.ID ,'ContentPlaceHolder1_txtUOTP')
email_otp=str(input("Enter the otp  here ... : "))
# while True:
    # if email_input is not None:
email_input.send_keys(email_otp)
        # break
driver.find_element(By.ID ,'ContentPlaceHolder1_btnValidate').click()
driver.maximize_window()
# hiting the link of BookKeeperView.aspx..........
time.sleep(6)
driver.get('https://stockexcel.com/BookKeeperView.aspx')
soup=BeautifulSoup(driver.page_source,'html.parser')
table=soup.find('table',{'id':'ContentPlaceHolder1_gvDebitNote'}).find_all('tr')
time.sleep(3)

# for i  in range (Start_range,End_range):
#for i  in range (1,len(table)):
for i  in range (3,len(table)-1,2):
# for i in range (3,10,2):
    print(table,"ttttttttttttttttttttttttttttttttt: table ")
    all_anchor_tag=table[i].find('td').find('a')
   
    driver.get("https://stockexcel.com/"+all_anchor_tag.get('href'))
    # driver.get('https://stockexcel.com/DNAdjustView.aspx?dn_id=92972')
    # ContentPlaceHolder1_gvDNLedger
    time.sleep(3)
    soup=BeautifulSoup(driver.page_source,'html.parser')
    
    
    table1=soup.find('table',{'id':'ContentPlaceHolder1_gvDNLedger'}).find_all('tr')
      # if company_name.text.upper() in comany_list:
     ###############  opening the windows of tally ##################
    def windowEnumHandler(hwnd, top_windows):
        top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))
    
    def bringToFront(window_name):
        top_windows = []
        win32gui.EnumWindows(windowEnumHandler, top_windows)
        for i in top_windows:
            # print(i[1])
            if window_name.lower() in i[1].lower():
                # print("found", window_name)
                win32gui.ShowWindow(i[0], win32con.SW_SHOWNORMAL)
                win32gui.SetForegroundWindow(i[0])
                break
    winname = "Tally.ERP 9 (Remote)"
    bringToFront(winname)
    
    time.sleep(3)
            
        # print(table1[i].find_all('td')[1].text.strip())
    pg.press("enter", presses=2)
     
    pg.write(table1[1].find_all('td')[1].text.strip())
    pg.hotkey('ctrl', 'a')
    
    # pg.press("enter", presses=1)
    pg.write(table1[1].find_all('td')[3].text.strip())
    pg.press("enter", presses=1)
    
    pg.write("n")
    pg.press("enter", presses=1)
    Debit_number=soup.find('span',{'id':'ContentPlaceHolder1_lblDNNo'})
    
    pg.write(Debit_number.text.strip())
    pg.press("enter", presses=4)
    for i in range(2,len(table1)-1):
        if "TDS" not in table1[i].find_all('td')[1].text.strip() :
            pg.write(table1[i].find_all('td')[1].text.strip())
            #print("nyyyyyyyyyyyyyyyyyyyyyyyyyy" , table1[i].find_all('td')[1].text.strip())
            
            pg.press("enter", presses=1)
            
            pg.write(table1[i].find_all('td')[4].text.strip())
            time.sleep(1)
            pg.press("enter", presses=2)
            time.sleep(1)
            pg.write(table1[i].find_all('td')[2].text.strip())
            time.sleep(1)
            # print(table1[i].find_all('td')[0].text)
            if i+1==len(table1)-1:
                
                    pg.press("enter", presses=3)
            else:
                pg.press("enter", presses=4)
                
                
        else:
            pg.write(table1[i].find_all('td')[1].text.strip())
            
            pg.press("enter", presses=1)
            time.sleep(1)
            
            pg.write(table1[i].find_all('td')[4].text.strip())
            time.sleep(1)
            pg.press("enter", presses=1)
            #print("noooooooooooooooooooooooooooooooooooooooooooooooooo" , table1[i].find_all('td')[1].text.strip())
    time.sleep(1)
    
    pg.write(Debit_number.text.strip())
    time.sleep(1)
    pg.press("enter", presses=2)
        
        
        
    time.sleep(1)
        
        
        
        