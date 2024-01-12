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
from selenium.webdriver.common.alert import Alert
from datetime import date
import pyautogui as pg
import time
import re
import win32con
import win32gui   
# https://www.awesomescreenshot.com/video/10149675?key=975ba5022f7eb37bc4e78d8886f3a7de
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
driver.get('https://stockexcel.com/PayCartNewly.aspx')
time.sleep(2)
soup=BeautifulSoup(driver.page_source,'html.parser')
table=soup.find_all('table',{'class':'text_style_custom'})[1].find_all('tr',{'style':'color:#000066;'})
# print(table,"tttttttttttttttttttttttttttt")
for i  in range (0,len(table)):
# for i in range(1, 2):


    all_anchor_tag="https://stockexcel.com/"+table[i].find_all('a')[1].get('href')

    driver.get(all_anchor_tag)
    time.sleep(3)
    soup=BeautifulSoup(driver.page_source,'html.parser')
    company_name=soup.find('span',{'id':'ContentPlaceHolder1_lblSupplierName'})
    invoice_no=soup.find('span',{'id':'ContentPlaceHolder1_lblBillNo'})
    Date=soup.find('span',{'id':'ContentPlaceHolder1_lblBillDate'})
    invoice_ammount=soup.find('span',{'id':'ContentPlaceHolder1_lblBillAmount'})
    narration=soup.find('span',{'id':'ContentPlaceHolder1_lblSRNO'})
    paid_ammaunt=soup.find('span',{'id':'ContentPlaceHolder1_lblBillPassedAmt_sel'})
    adv_invoice_amount=soup.find('span',{'id':'ContentPlaceHolder1_lblReqAmt'})
    adv_tds_on_good=soup.find('span',{'id':'ContentPlaceHolder1_lblTDSRate'})
    adv_tds_amount=soup.find('span',{'id':'ContentPlaceHolder1_lblTDSAmt'})
    ContentPlaceHolder1_gvDNAdjust=soup.find('table',{'id':'ContentPlaceHolder1_gvDNAdjust'})
    
    """
    opening tally 
    """
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
     # case of advance paymnet ...   
    if ContentPlaceHolder1_gvDNAdjust:
            print("paymet with debit note .....")   
            
            pg.write(company_name.text.strip())
            time.sleep(1)
            pg.press("enter", presses=1)
            time.sleep(1)
            
            pg.write(paid_ammaunt.text.strip().replace("Rs.",'').split(".")[0])
            time.sleep(1)
            pg.press("enter")
            
            pg.write("ag")
            time.sleep(1)
            pg.press("enter",presses=1)
            time.sleep(1)
            pg.write(invoice_no.text.strip()+" "+"TDS")
            
            
    elif "ADV" not in narration.text.strip():
        time.sleep(3)
        total_paid_bill=soup.find('table',{'id':'ContentPlaceHolder1_gvLedgerBillPros'}).find_all('tr')[-1].find_all('td')[-1]
  
        table1=soup.find('table',{'id':'ContentPlaceHolder1_gvLedgerBillPros'}).find_all('tr')
        tds_check_list=[]
        for index in range(1,len(table1)-1):
            
            if "TDS" in table1[index].find_all('td')[0].text.strip():
                tds_check_list.append(table1[index].find_all('td')[0].text.strip())
                
        if len(tds_check_list)>0:
            print("tds in without advvance ........")
            pg.write(company_name.text.strip())
            time.sleep(1)
            pg.press("enter", presses=1)
            time.sleep(1)
            
            pg.write(paid_ammaunt.text.strip().replace("Rs.",'').split(".")[0])
            time.sleep(1)
            pg.press("enter")
            
            pg.write("ag")
            time.sleep(1)
            pg.press("enter",presses=1)
            time.sleep(1)
            pg.write(invoice_no.text.strip()+" "+"TDS")
         
            time.sleep(1)
            
            pg.press("enter" ,presses=3)
            pg.write("ag")
            time.sleep(1)
            pg.press("enter",presses=1)
            time.sleep(1)
            pg.write(invoice_no.text.strip())
            time.sleep(1)
            pg.press("enter",presses=1)
            pg.write(invoice_ammount.text.strip())
            time.sleep(1)
            pg.press("enter",presses=3)
            pg.write("HDFC Bank - 9809")
            time.sleep(1)
           
            pg.press("enter", presses=2)
            pg.write("e-Fund Transfer")
            time.sleep(1)
            
            pg.press("enter", presses=7)
            time.sleep(1)
            
            pg.write("KIL-"+narration.text.strip())
            time.sleep(1)
            pg.press("enter", presses=2)
          
    
            
        else:
            print("tds not in thins ..........................")
            pg.write(company_name.text.strip())
            time.sleep(1)
            pg.press("enter", presses=1)
            time.sleep(1)
            
            pg.write(paid_ammaunt.text.strip().replace("Rs.",'').split(".")[0])
            time.sleep(1)
            pg.press("enter")
            
            pg.write("ag")
            time.sleep(1)
            pg.press("enter",presses=1)
            time.sleep(1)
            pg.write(invoice_no.text.strip())
         
            time.sleep(1)
            
            pg.press("enter" ,presses=4)
            pg.write("HDFC Bank - 9809")
            time.sleep(1)
           
            pg.press("enter", presses=2)
            pg.write("e-Fund Transfer")
            time.sleep(1)
            
            pg.press("enter", presses=7)
            time.sleep(1)
            
            pg.write("KIL-"+narration.text.strip())
            time.sleep(1)
            
        
            pg.press("enter", presses=2)
            

        
    

  
    
    else:
        if adv_tds_on_good :
            print('adv paymet with TDS ..........')
            time.sleep(2)
            pg.write(company_name.text.strip().replace("Limited",""))
            time.sleep(2)
            pg.press("enter", presses=1)
            time.sleep(1)
            
            pg.write(adv_invoice_amount.text.strip().replace('Rs.',''))
            time.sleep(1)
            
            pg.press("enter", presses=1)
            pg.write("n")
            time.sleep(1)
            pg.press("enter", presses=1)
            pg.write(narration.text.strip())
            time.sleep(0.5)
            pg.press("enter", presses=1)
            time.sleep(0.5)
            pg.press("enter", presses=1)
            time.sleep(0.5)
            pg.press("enter", presses=1)
            time.sleep(0.5)
            pg.press("enter", presses=1)
            
            time.sleep(0.5)
            pg.press("enter", presses=1)
            pg.write(adv_tds_on_good.text.strip().replace("[","").replace("]",""))
            pg.press('enter')
            pg.write(adv_tds_amount.text.strip())
            time.sleep(0.5)
            pg.press('enter')
            time.sleep(0.5)
            pg.press('enter')
            time.sleep(0.5)
            pg.write("HDFC Bank - 9809")
            time.sleep(1)
          
            pg.press("enter", presses=2)
            pg.write("e-Fund Transfer")
            time.sleep(1)
            
            pg.press("enter", presses=7)
            time.sleep(1)
            
            pg.write("KIL-"+narration.text.strip())
            time.sleep(1)
            # pg.press("enter", presses=2)
            # time.sleep(100)
            
            pg.write("KIL-"+narration.text.strip())
            pg.press("enter", presses=2)
            
        else:
        
            print("advancve payment without TDS  ...............")
            time.sleep(2)
            pg.write(company_name.text.strip())
            time.sleep(2)
            pg.press("enter", presses=1)
            time.sleep(1)
            
            pg.write(adv_invoice_amount.text.strip().replace('Rs.',''))
            time.sleep(1)
            
            pg.press("enter", presses=1)
            pg.write("n")
            time.sleep(1)
            pg.press("enter", presses=1)
            pg.write(narration.text.strip())
            time.sleep(0.5)
            pg.press("enter", presses=1)
            time.sleep(0.5)
            pg.press("enter", presses=1)
            time.sleep(0.5)
            pg.press("enter", presses=1)
            time.sleep(0.5)
            pg.press("enter", presses=1)
            
            time.sleep(0.5)
            pg.press("enter", presses=1)
            pg.write("HDFC Bank - 9809")
            time.sleep(1)
          
            pg.press("enter", presses=2)
            pg.write("e-Fund Transfer")
            time.sleep(1)
            
            pg.press("enter", presses=7)
            time.sleep(1)
            
            pg.write("KIL-"+narration.text.strip())
            time.sleep(1)
            # pg.press("enter", presses=2)
            
            pg.write("KIL-"+narration.text.strip())
            time.sleep(1)
           
            pg.press("enter", presses=2)
            