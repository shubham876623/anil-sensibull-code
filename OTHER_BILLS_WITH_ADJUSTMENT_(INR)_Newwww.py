import requests
import pandas as pd
from io import BytesIO
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
# from emailotp import get_otp
from selenium.webdriver.common.alert import Alert
from datetime import date
import pyautogui as pg
import time
import re
import win32con
import win32gui                  
discount_list=["Hindoostan Mills Limited" ,"SREE Karpagambal MILLS LTD" ,"NAHAR INDUSTRIAL ENT.LTD","AJANTA Universal LTD"]                                                                                                                                 
comany_list=["CALCUTTA EXPRESS TRANSPORT CO.","A K ENTERPRISE","INTERTEK INDIA PVT LTD","PRYM FASHION ASIA PACIFIC LIMITED","PRYM FASHION GMBH","HINDOOSTAN MILLS LIMITED","DHL EXPRESS IGST","SPAN PUBLICITY","AJANTA TEXTILES","BKS TEXTILES PVT LTD","PETEX","PRYM FASHION ASIA PACIFIC LIMLTD","SHILPA ENTERPRISES","DIPAK TRADING COMPANY","JAIN NARROW FABRICS","THREADS INDIA LTD","ANSUN MULTITECH INDIA LTD","AMFINO TEX INDIA PVT LTD","CASH","YKK INDIA PVT LTD","VARTIKA PACKAGING P LTD","ANTARCTICA LIMITED","DEEYA INTERNATIONAL","KALPATARU PVT LTD","PECKEL","SATTIK PACKAGING PVT LTD","SREE KARPAGAMBAL MILLS LTD","JUKI INDIA PRIVATE LIMITED"]

# driver=webdriver.Chrome(ChromeDriverManager(version='114.0.5735.90').install())
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
email_input.send_keys(email_otp)
driver.find_element(By.ID ,'ContentPlaceHolder1_btnValidate').click()
driver.maximize_window()
# hiting the link of BookKeeperView.aspx..........
time.sleep(3)
driver.get('https://stockexcel.com/BookKeeperView.aspx')
time.sleep(5)
soup=BeautifulSoup(driver.page_source,'html.parser')
table=soup.find('table',{'id':'ContentPlaceHolder1_gvOtherBillsWA'}).find_all('tr')
time.sleep(3)
for i  in range (1,len(table)-1,2):
#for i in range(1, 19,2):


    all_anchor_tag=table[i].find_all('td')[1].find('a')
    # print("https://stockexcel.com/"+all_anchor_tag.get('href'))
    driver.get("https://stockexcel.com/"+all_anchor_tag.get('href'))
    # driver.get('https://stockexcel.com/BillProcess.aspx?bpd=52738&bk=y')
    time.sleep(3)
    soup=BeautifulSoup(driver.page_source,'html.parser')
    company_name=soup.find('span',{'id':'ContentPlaceHolder1_lblSupplierName'})
    invoice_no=soup.find('span',{'id':'ContentPlaceHolder1_lblBillNo'})
    Date=soup.find('span',{'id':'ContentPlaceHolder1_lblBillDate'})
    invoice_ammount=soup.find('span',{'id':'ContentPlaceHolder1_lblBillAmount'})
    narration=soup.find('span',{'id':'ContentPlaceHolder1_lblSRNO'})
    paid_ammaunt=soup.find('span',{'id':'ContentPlaceHolder1_lblBillPassedAmt_sel'})
    # currency=soup.find('span',{'id':''})
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
    winname = "Tally.ERP 9"
    bringToFront(winname)
    
    time.sleep(3)
    table1=soup.find('table',{'id':'ContentPlaceHolder1_gvLedgerBillPros'}).find_all('tr')
            
            
            
    # entering 	Bill Ledgers  table :
    list_a=[]
    total_paid_bill=soup.find('table',{'id':'ContentPlaceHolder1_gvLedgerBillPros'}).find_all('tr')[-1].find_all('td')[-1]
    
    for i in range(1,len(table1)-1):
        
        if "TDS" not in table1[i].find_all('td')[0].text.strip() :
            if i>1:

                if "-" in table1[i].find_all('td')[2].text.strip():
                
                    pg.write("Cr")
                    pg.press("enter")        
                    
                else:
                    pg.write("Dr")
                    pg.press("enter")        
            pg.write(table1[i].find_all('td')[0].text)
            time.sleep(1)
            # hjhhghgyftyrtrtr5hgyugyukfyyuf
            pg.press("enter", presses=1)
            time.sleep(1)
            
            if "-" in table1[i].find_all('td')[2].text.strip():
                list_a.append(table1[i].find_all('td')[2].text.strip()[1:])
            else: 
                list_a.append(table1[i].find_all('td')[2].text.strip())
                
            
            pg.write(list_a[-1])
            # time.sleep(1)
            # pg.write("600")
            time.sleep(1)
            pg.press("enter", presses=2)
            time.sleep(1)

        
            pg.write(table1[i].find_all('td')[1].text.strip())
            time.sleep(1)
            print(table1[i].find_all('td')[0].text)
            # print(i,len(table1))
            pg.press("enter", presses=4)
            pg.press("backspace", presses=1)
        else:
            # list_a=[]
            # print("YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYyyyyyyyyyyyyyyyyyyyyyyyyy")
            # pg.press("enter", presses=1)
            
            if "-" in table1[i].find_all('td')[2].text.strip():
                
                    pg.write("Cr")
                    pg.press("enter")        
                    
            else:
                pg.write("Dr")
                pg.press("enter")        
            pg.write(table1[i].find_all('td')[0].text)
            time.sleep(1)
            # time.sleep(1)
            # hjhhghgyftyrtrtr5hgyugyukfyyuf
            pg.press("enter", presses=1)
            
            time.sleep(1)
            if "-" in table1[i].find_all('td')[2].text.strip():
                list_a.append(table1[i].find_all('td')[2].text.strip()[1:])
            else: 
                list_a.append(table1[i].find_all('td')[2].text.strip())
            pg.write(list_a[-1])
                
            pg.press("enter", presses=1)
                
            
    # entering bill-wise detail  ..
    
    if (invoice_ammount.text).strip()==(total_paid_bill.text).strip():
        print("YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYyyyyyyyyyyyyyyyyyyyyyyyyy")
        pg.press("enter")        
                
        pg.write(company_name.text.strip())  # entering the supplier name 
        pg.press("enter", presses=1)
        pg.write(total_paid_bill.text.strip()) # entering the paid amout 
        pg.press("enter", presses=1)
        pg.write("N")
        pg.press("enter", presses=1)
        
        pg.write(invoice_no.text.strip())     # entering the invoice_number 
        time.sleep(2)
    
        pg.press("enter", presses=2)
        pg.write(invoice_ammount.text.strip())    # entering the invoice_amount
        time.sleep(2)
        pg.press("enter", presses=2)
        time.sleep(2)
        
        pg.write("KIL-"+narration.text.strip()) 
        pg.press("enter", presses=2)
        
    else:
    
        pg.press("enter")        
        time.sleep(1)      
        pg.write(company_name.text.strip())  # entering the supplier name 
        time.sleep(1)
        pg.press("enter", presses=1)
        time.sleep(1)
        pg.write(total_paid_bill.text.strip()) # entering the paid amout 
        time.sleep(1)
        pg.press("enter", presses=1)
        time.sleep(1)
        pg.write("N")                         # press N to open the new refrence 
        time.sleep(1)
        pg.press("enter", presses=1)
        time.sleep(1)
        pg.write(invoice_no.text.strip())     # entering the invoice_number 
        time.sleep(2)
    
        pg.press("enter", presses=2)
        time.sleep(1)
        pg.write(invoice_ammount.text.strip())    # entering the invoice_amount
        time.sleep(2)
    
        pg.press("enter", presses=2)
        time.sleep(1)
        pg.write("N")                       # opening new refrence here ..
        pg.press("enter", presses=1)
        
        
        time.sleep(2)
    
        pg.write(invoice_no.text.strip()+"TDS")    # en invoice number with tds 
        time.sleep(2)
    
        pg.press("enter", presses=4)
        time.sleep(2)
        
        pg.write("KIL-"+narration.text.strip()) 
        time.sleep(2)
        pg.press("enter", presses=2)
        
                 
         
    
            
            
            
            
            
            
    # else:
    #     print("no comapnty name ",company_name.text.upper())

time.sleep(100)
    

