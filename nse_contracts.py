from selenium import webdriver
from selenium.webdriver.support import ui
from selenium.webdriver.support.ui import WebDriverWait
from datetime import date, timedelta
import datetime
import pandas as pd
import os
from openpyxl import load_workbook

import time

#columns=["Symbol","Date1","Date2","Chg in OI","% Chg in OI","Volume","Futures","Options (Notional)","Total","Options(Premium)","Underlyning Value"]
#df=pd.DataFrame(columns=columns);

def page_is_loaded(driver):
    # this function waits for the page to load before performing further actions.
    return driver.find_element_by_tag_name("body") != None



def contracts(date):
    print("NSE CONTRACTS")
    driver = webdriver.Chrome(r'E:\chromedriver.exe')

    #driver.get("https://www.zaubacorp.com/")
    driver.get("https://www.nseindia.com/products/content/equities/equities/oi_spurts.htm")


    time.sleep(4);

    wait = ui.WebDriverWait(driver, 10)
    wait.until(page_is_loaded)

    ##END OF UNDERLYNING

    button=driver.find_element_by_xpath('//*[@id="tab8"]')

    button.click()
    time.sleep(5)
    df1=pd.DataFrame()

    columns1=["Instrument","Symbol","Expiry","Strike Price","Type","LTP","Prev.Close","%Change in LTP",date+" OI","Jan24,2018 OI","OI Change","Volume in contracts","TurnOver in crores","Premium Turnover in crores","Underlyning Value","Type of OI Spurts","Current Business Date","Previous Business Date"]
    #s=['Rise in OI-Rise in Price','Rise in OI-Slide in Price','Slide in OI-Rise in Price','Slide in OI-Slide in Price']

    k=0

    ul=driver.find_elements_by_xpath('//*[@id="replacetext"]/div/ul')

    ul=['//*[@id="replacetext"]/div/ul/li[2]','//*[@id="replacetext"]/div/ul/li[3]','//*[@id="replacetext"]/div/ul/li[4]']

    time.sleep(5)
    row_count=len(driver.find_elements_by_xpath('//*[@id="replacetext"]/table/tbody/tr'))
    col_count=len(driver.find_elements_by_xpath('//*[@id="replacetext"]/table/tbody/tr[2]/td'))
    print (row_count)
    print (col_count)

    i=2
    d1={}
    first_str='//*[@id="replacetext"]/table/tbody/tr['
    second_str=']/td['
    third_str=']'


    print ("FIRST ")
    while i<=row_count:
        j=1
        while j<=col_count:
            final_str=first_str+str(i)+second_str+str(j)+third_str
            d1[columns1[j-1]]=driver.find_element_by_xpath(final_str).text
            j=j+1
        d1[columns1[j-1]]="Rise in OI-Rise in Price"
        j=j+1
        d1[columns1[j-1]]=date
        j=j+1
        d1[columns1[j-1]]='Jan24,2018'
        i=i+1
        df1=df1.append(d1,ignore_index=True)

    button=driver.find_element_by_xpath('//*[@id="riseinOIslideinPrice"]')
    button.click()

    time.sleep(5)
    print ("SECOND ")

    row_count=len(driver.find_elements_by_xpath('//*[@id="replacetext"]/table/tbody/tr'))
    col_count=len(driver.find_elements_by_xpath('//*[@id="replacetext"]/table/tbody/tr[2]/td'))

    i=2
    while i<=row_count:
        j=1
        while j<=col_count:
            final_str=first_str+str(i)+second_str+str(j)+third_str
            d1[columns1[j-1]]=driver.find_element_by_xpath(final_str).text
            j=j+1
        d1[columns1[j-1]]="Rise in OI-Slide in Price"
        j=j+1
        d1[columns1[j-1]]=date
        j=j+1
        d1[columns1[j-1]]='Jan24,2018'
        i=i+1
        df1=df1.append(d1,ignore_index=True)

    button=driver.find_element_by_xpath('//*[@id="slideinOIriseinPrice"]')
    button.click()
    time.sleep(5)
    print ("third")

    row_count=len(driver.find_elements_by_xpath('//*[@id="replacetext"]/table/tbody/tr'))
    col_count=len(driver.find_elements_by_xpath('//*[@id="replacetext"]/table/tbody/tr[2]/td'))

    i=2
    while i<=row_count:
        j=1
        while j<=col_count:
            final_str=first_str+str(i)+second_str+str(j)+third_str
            d1[columns1[j-1]]=driver.find_element_by_xpath(final_str).text
            j=j+1
        d1[columns1[j-1]]="Slide in OI-Rise in Price"
        j=j+1
        d1[columns1[j-1]]=date
        j=j+1
        d1[columns1[j-1]]='Jan24,2018'
        i=i+1
        df1=df1.append(d1,ignore_index=True)

    button=driver.find_element_by_xpath('//*[@id="slideinOIslideinPrice"]')
    button.click()
    time.sleep(5)
    print ("fourth")

    row_count=len(driver.find_elements_by_xpath('//*[@id="replacetext"]/table/tbody/tr'))
    col_count=len(driver.find_elements_by_xpath('//*[@id="replacetext"]/table/tbody/tr[2]/td'))

    i=2
    while i<=row_count:
        j=1
        while j<=col_count:
            final_str=first_str+str(i)+second_str+str(j)+third_str
            d1[columns1[j-1]]=driver.find_element_by_xpath(final_str).text
            j=j+1
        d1[columns1[j-1]]="Slide in OI-Slide in Price"
        j=j+1
        d1[columns1[j-1]]=date
        j=j+1
        d1[columns1[j-1]]='Jan24,2018'
        i=i+1
        df1=df1.append(d1,ignore_index=True)

    f = open(date + ".xlsx")
    st = os.path.realpath(f.name)
    path = st
    book = load_workbook(path)
    writer = pd.ExcelWriter(path, engine='openpyxl')
    writer.book = book
    df1.to_excel(writer,sheet_name="NSE CONTRACTS",columns=["Instrument","Symbol","Expiry","Strike Price","Type","LTP","Prev.Close","%Change in LTP",date+" OI","Jan24,2018 OI","OI Change","Volume in contracts","TurnOver in crores","Premium Turnover in crores","Underlyning Value","Type of OI Spurts","Current Business Date","Previous Business Date"])
    writer.save()
    writer.close()
    driver.close()
