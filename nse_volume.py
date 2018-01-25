from selenium import webdriver
from selenium.webdriver.support import ui
from selenium.webdriver.support.ui import WebDriverWait
import os
from openpyxl import load_workbook
import pandas as pd
 

import time

def page_is_loaded(driver):
    # this function waits for the page to load before performing further actions.
    return driver.find_element_by_tag_name("body") != None


def volume(date):
    print("VOLUME")
    driver = webdriver.Chrome(r'E:\chromedriver.exe')

    columns=["Symbol","Security","Volume","Avg.Volume for 1 Week","Change(No.of time) for 1 Week","Avg.Volume for 2 Weeks","Change(No.of times) for 2 Weeks","LTP for Today","%Chng for Today","TurnOver for Today","Current Business Date"]
    df=pd.DataFrame()
    driver.get("https://www.nseindia.com/live_market/dynaContent/live_watch/volume_spurts.htm")
    time.sleep(4)
    wait = ui.WebDriverWait(driver, 10)
    wait.until(page_is_loaded)


    #driver.get("https://www.zaubacorp.com/")

    row_count=len(driver.find_elements_by_xpath('//*[@id="dataTable"]/tbody/tr'))
    col_count=len(driver.find_elements_by_xpath('//*[@id="dataTable"]/tbody/tr[3]/td'))
    i=3
    d={}
    print (row_count)
    print (col_count)

    first_str='//*[@id="dataTable"]/tbody/tr['
    second_str=']/td['
    third_str=']'

    while i<=row_count:
        j=1
        while j<=col_count:
            final_str=first_str+str(i)+second_str+str(j)+third_str
            d[columns[j-1]]=driver.find_element_by_xpath(final_str).text
            j=j+1
        d[columns[j-1]]=date
        i=i+1
        df=df.append(d,ignore_index=True)
        print (d)

    f = open(date + ".xlsx")
    path = os.path.realpath(f.name)
    book = load_workbook(path)
    writer = pd.ExcelWriter(path, engine='openpyxl')
    writer.book = book
    df.to_excel(writer,sheet_name="NSE VOLUME",columns=["Symbol","Security","Volume","Avg.Volume for 1 Week","Change(No.of time) for 1 Week","Avg.Volume for 2 Weeks","Change(No.of times) for 2 Weeks","LTP for Today","%Chng for Today","TurnOver for Today","Current Business Date"])
    writer.save()
    writer.close()
    driver.close()
