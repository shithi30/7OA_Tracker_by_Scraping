#!/usr/bin/env python
# coding: utf-8

## import
import pandas as pd
import duckdb
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import random
from fuzzywuzzy import process
import win32com.client
from googleapiclient.discovery import build
from google.oauth2 import service_account
import time

## particulars
brands = ['Camay', 'Pears', 'Signal Paste', 'Clear Shampoo', 'Simple Fac', 'Simple Mask', 'Pepsodent', 'Brylcreem', 'St. Ives', 'St.Ives', 'Sunsilk', 'Sun Silk', 'Lux', 'Ponds', "Pond's", 'Closeup', 'Close Up', 'Dove', 'Clinic Plus', 'Tresemme', 'Tresemmé', 'Glow Lovely', 'Fair Lovely', 'Glow Handsome', 'Axe Body', 'Lifebuoy', 'Vaseline']
keywords = [
    'Camay soap', 'Camay bar', 'Camay saban',
    'Pears soap', 'Pears soap bar', 'Pears saban',
    'Signal paste', 'Signal toothpaste',
    'Clear dandruff shampoo', 'Clear hair', 'Clear shampoo',
    'Simple moisturizer', 'Simple facial', 
    'Pepso dent', 'Pepso dent paste',
    'Bryl Creem', 'Bralcream', 'Bryl cream',
    'San Ive', 'Saint Eve', 'San Eve',
    'Sun silk hair', 'Sunsilk shampoo', 'Sunsilk condition',
    'Lux soap', 'Lux saban', 'Lux bar', 'Lux body', 'Lux body wash',
    'Ponds cream', 'Ponds lotion', 'Pond cream',
    'Close up mouth', 'Close up paste', 'Close up tooth',
    'Dove soap bar', 'Dove saban', 'Dove cream',
    'Clinic Plus shampoo', 'Clinic Plus hair', 'Clinic Plus dandruff', 
    'Treseme shampoo', 'Tressem hair', 'Tresme shampoo',
    'Glow & lovely', 'Glow and lovely', 'Fair & lovely', 'Fair and lovely',
    'Glow & handsome', 'Glow and handsome', 'Fair & handsome', 'Fair and handsome',
    'Axe deo', 'Axe perfum', 'Axe spray', 'Axe body spray',
    'Lifeboy', 'Life boy', 'Life buoy',
    'Vaslin', 'Veslin', 'Vase lin'
]

## preference
options = webdriver.ChromeOptions()
options.add_argument('ignore-certificate-errors')

## open window
driver = webdriver.Chrome(options = options)
driver.maximize_window()

## comms
def comms(platform, sos_0_list):
    
    # expressions, ref: https://www.w3schools.com/charsets/ref_emoji.asp
    emos = [9888, 9940, 9889, 128128, 127875, 128030, 128110, 128073
        # exclaim, minus, thunder, skull, pumpkim, bug, police, indicate
    ]
    
    # summary
    sos = str(sos_0_list)[1:-1]
    sos = "&#" + str(random.choice(emos)) + " Keywords with 0 SoS: <i>" + sos.replace("'", "") + '''</i><br><br>View detailed results <a href="https://docs.google.com/spreadsheets/d/1gkLRp59RyRw4UFds0-nNQhhWOaS4VFxtJ_Hgwg2x2A0/edit#gid=935743274">here</a>.<br>''' if len(sos)>2 else ""

    # email
    ol = win32com.client.Dispatch("outlook.application")
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)

    # Teams
    newmail.Subject = platform + " 0 SoS"
    newmail.CC = "Eagle Eye - Alerts <93f21d6e.Unilever.onmicrosoft.com@emea.teams.ms>"
    newmail.HTMLbody = sos + "<br>"
    if len(sos) > 0: newmail.Send()

## Shajgoj

# accumulators
start_time = time.time()
df_acc_shaj = pd.DataFrame()
sos_0_keywords = []

# url
driver.get("https://shop.shajgoj.com/shop/")

# keyword
for k in keywords:
    print("Scraping for keyword: " + k)
    
    # search
    elem = driver.find_element(By.CLASS_NAME, "ais-search-box--input")
    elem.clear()
    elem.send_keys(k + "\n")

    # soup
    time.sleep(5)
    soup = BeautifulSoup(driver.page_source, 'html.parser').find_all("div", attrs={"class": "ais-infinite-hits--item ais-hits--item"})

    # scrape
    skus = []
    for s in soup:
        try: val = s.find("a", attrs={"class": "product_title"}).get_text()
        except: val = None
        skus.append(val)

    # accumulate
    df = pd.DataFrame()
    df['basepack'] = skus
    df['keyword'] = k
    if_ubl = []

    # Unilever
    sku_count = len(skus)
    b = process.extractOne(k.lower(), [b.lower() for b in brands])[0] 
    for i in range(0, sku_count):
        if_ubl.append(None)
        bb = b.split()
        if len(bb) == 1: bb.append('')
        if bb[0] + ' ' in skus[i].lower() and bb[1] in skus[i].lower(): if_ubl[i] = 1
    df['brand_unilever'] = if_ubl    

    # record
    df = duckdb.query('''select 'Shajgoj' platform, keyword, basepack, left(now()::text, 19) report_time from df where brand_unilever = 1''').df()
    if df.shape[0] == 0: sos_0_keywords.append(k)
    df_acc_shaj = df_acc_shaj.append(df)
    
# comms
comms("Shajgoj", sos_0_keywords)
    
# stats
display(df_acc_shaj)
print("0 SoS search terms:\n" + str(sos_0_keywords))
print("\nElapsed time to report (mins): " + str(round((time.time() - start_time) / 60.00, 2)))

## OHSOGO

# accumulators
start_time = time.time()
df_acc_osgo = pd.DataFrame()
sos_0_keywords = []

# keyword
for k in keywords:
    print("Scraping for keyword: " + k)
    driver.get("https://ohsogo.com/search?q=" + k)

    # soup
    time.sleep(5)
    soup = BeautifulSoup(driver.page_source, "html.parser").find_all("div", attrs={"class": "zv-title--kF9LU"})
    
    # scrape
    skus = []
    for s in soup:
        try: val = s.get_text()
        except: val = None
        skus.append(val)

    # accumulate
    df = pd.DataFrame()
    df['basepack'] = skus
    df['keyword'] = k
    if_ubl = []

    # Unilever
    sku_count = len(skus)
    b = process.extractOne(k.lower(), [b.lower() for b in brands])[0] 
    for i in range(0, sku_count):
        if_ubl.append(None)
        bb = b.split()
        if len(bb) == 1: bb.append('')
        if bb[0] + ' ' in skus[i].lower() and bb[1] in skus[i].lower(): if_ubl[i] = 1
    df['brand_unilever'] = if_ubl    

    # record
    df = duckdb.query('''select 'OHSOGO' platform, keyword, basepack, left(now()::text, 19) report_time from df where brand_unilever = 1''').df()
    if df.shape[0] == 0: sos_0_keywords.append(k)
    df_acc_osgo = df_acc_osgo.append(df)
    
# comms
comms("OHSOGO", sos_0_keywords)
    
# stats
display(df_acc_osgo)
print("0 SoS search terms:\n" + str(sos_0_keywords))
print("\nElapsed time to report (mins): " + str(round((time.time() - start_time) / 60.00, 2)))

## Chaldal

# accumulators
start_time = time.time()
df_acc_cldl = pd.DataFrame()
sos_0_keywords = []

# keyword
for k in keywords:
    print("Scraping for keyword: " + k)
    driver.get("https://chaldal.com/search/" + k)

    # soup
    time.sleep(5)
    soup = BeautifulSoup(driver.page_source, "html.parser").find_all("div", attrs={"class": "product"})
    
    # scrape
    skus = []
    for s in soup:
        try: val = s.find("div", attrs={"class": "name"}).get_text()
        except: val = None
        skus.append(val)

    # accumulate
    df = pd.DataFrame()
    df['basepack'] = skus
    df['keyword'] = k
    if_ubl = []

    # Unilever
    sku_count = len(skus)
    b = process.extractOne(k.lower(), [b.lower() for b in brands])[0] 
    for i in range(0, sku_count):
        if_ubl.append(None)
        bb = b.split()
        if len(bb) == 1: bb.append('')
        if bb[0] + ' ' in skus[i].lower() and bb[1] in skus[i].lower(): if_ubl[i] = 1
    df['brand_unilever'] = if_ubl    

    # record
    df = duckdb.query('''select 'Chaldal' platform, keyword, basepack, left(now()::text, 19) report_time from df where brand_unilever = 1''').df()
    if df.shape[0] == 0: sos_0_keywords.append(k)
    df_acc_cldl = df_acc_cldl.append(df)
    
# comms
comms("Chaldal", sos_0_keywords)
    
# stats
display(df_acc_cldl)
print("0 SoS search terms:\n" + str(sos_0_keywords))
print("\nElapsed time to report (mins): " + str(round((time.time() - start_time) / 60.00, 2)))

## Pandamart

# accumulators
start_time = time.time()
df_acc_pmrt = pd.DataFrame()
sos_0_keywords = []

# url
driver.get("https://www.foodpanda.com.bd/darkstore/w2lx/pandamart-gulshan-w2lx")

# keyword
for k in keywords:
    print("Scraping for keyword: " + k)
    elem = driver.find_element(By.XPATH, '//*[@id="groceries-menu-react-root"]/div/div/div[2]/div/section/div[3]/div/div/div/div/div[1]/input')
    elem.send_keys(k + "\n")

    # soup
    time.sleep(5)
    soup = BeautifulSoup(driver.page_source, "html.parser").find_all('div', attrs={'class', 'box-flex product-card-attributes'})
    
    # scrape
    skus = []
    for s in soup:
        try: val = s.find('p', attrs={'class', 'product-card-name'}).get_text()
        except: val = None
        skus.append(val)

    # accumulate
    df = pd.DataFrame()
    df['basepack'] = skus
    df['keyword'] = k
    if_ubl = []

    # Unilever
    sku_count = len(skus)
    b = process.extractOne(k.lower(), [b.lower() for b in brands])[0] 
    for i in range(0, sku_count):
        if_ubl.append(None)
        bb = b.split()
        if len(bb) == 1: bb.append('')
        if bb[0] + ' ' in skus[i].lower() and bb[1] in skus[i].lower(): if_ubl[i] = 1
    df['brand_unilever'] = if_ubl    

    # record
    df = duckdb.query('''select 'Pandamart' platform, keyword, basepack, left(now()::text, 19) report_time from df where brand_unilever = 1''').df()
    if df.shape[0] == 0: sos_0_keywords.append(k)
    df_acc_pmrt = df_acc_pmrt.append(df)
    
    # back
    driver.back()
    
# comms
comms("Pandamart", sos_0_keywords)
    
# stats
display(df_acc_pmrt)
print("0 SoS search terms:\n" + str(sos_0_keywords))
print("\nElapsed time to report (mins): " + str(round((time.time() - start_time) / 60.00, 2)))

## Daraz

# accumulators
start_time = time.time()
df_acc_daaz = pd.DataFrame()
sos_0_keywords = []

# url
driver.get('https://www.daraz.com.bd/')

# keyword
for k in keywords:
    print("Scraping for keyword: " + k)
    elem = driver.find_element(By.ID, "q")
    elem.send_keys(Keys.CONTROL + "a")
    elem.send_keys(Keys.DELETE)
    elem.send_keys(k + "\n")

    # soup
    time.sleep(5)
    soup = BeautifulSoup(driver.page_source, "html.parser").find_all("div", attrs={"class": "gridItem--Yd0sa"})
    
    # scrape
    skus = []
    for s in soup:
        try: val = s.find("div", attrs={"id": "id-title"}).get_text()
        except: val = None
        skus.append(val)

    # accumulate
    df = pd.DataFrame()
    df['basepack'] = skus
    df['keyword'] = k
    if_ubl = []

    # Unilever
    sku_count = len(skus)
    b = process.extractOne(k.lower(), [b.lower() for b in brands])[0] 
    for i in range(0, sku_count):
        if_ubl.append(None)
        bb = b.split()
        if len(bb) == 1: bb.append('')
        if bb[0] + ' ' in skus[i].lower() and bb[1] in skus[i].lower(): if_ubl[i] = 1
    df['brand_unilever'] = if_ubl    

    # record
    df = duckdb.query('''select 'Daraz' platform, keyword, basepack, left(now()::text, 19) report_time from df where brand_unilever = 1''').df()
    if df.shape[0] == 0: sos_0_keywords.append(k)
    df_acc_daaz = df_acc_daaz.append(df)
    
    # back
    driver.back()
    
# comms
comms("Daraz", sos_0_keywords)
    
# stats
display(df_acc_daaz)
print("0 SoS search terms:\n" + str(sos_0_keywords))
print("\nElapsed time to report (mins): " + str(round((time.time() - start_time) / 60.00, 2)))

## combine
qry = '''
select * from df_acc_shaj union all
select * from df_acc_osgo union all
select * from df_acc_cldl union all
select * from df_acc_pmrt union all
select * from df_acc_daaz
'''
df_acc = duckdb.query(qry).df()

## GSheet

# credentials
SERVICE_ACCOUNT_FILE = 'read-write-to-gsheet-apis-1-04f16c652b1e.json'
SAMPLE_SPREADSHEET_ID = '1gkLRp59RyRw4UFds0-nNQhhWOaS4VFxtJ_Hgwg2x2A0'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# APIs
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=creds)
sheet = service.spreadsheets()

# update
sheet.values().clear(spreadsheetId=SAMPLE_SPREADSHEET_ID, range='0 SoS').execute()
sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="0 SoS!A1", valueInputOption='USER_ENTERED', body={'values': [df_acc.columns.values.tolist()] + df_acc.values.tolist()}).execute()

## close window
driver.close()
