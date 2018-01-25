from selenium import webdriver
from selenium.webdriver.support import ui
from selenium.webdriver.support.ui import WebDriverWait
import pandas as pd
import time
import os
from openpyxl import load_workbook

def page_is_loaded(driver):
    # this function waits for the page to load before performing further actions.
    return driver.find_element_by_tag_name("body") != None




def underlying(cbd):
    print ("NSE UNDERLYING")
    driver = webdriver.Chrome(r'E:\chromedriver.exe')
    driver.get("https://www.nseindia.com/products/content/equities/equities/oi_spurts.htm")


    columns=["Symbol",cbd+" OI","Jan24 OI","Change in OI","% Change in OI","Volume in Contracts","Futures in Crores","Options(Notional) in Crores","Total in Crores","Options(Premium) in Crores","Underlyning Value","Current Business Date","Previous Business Date"]
    df=pd.DataFrame()

    time.sleep(4)

    wait = ui.WebDriverWait(driver, 10)
    wait.until(page_is_loaded)


    row_count=len(driver.find_elements_by_xpath('//*[@id="dataTable"]/tbody/tr'))
    col_count=len(driver.find_elements_by_xpath('//*[@id="dataTable"]/tbody/tr[3]/td'))

    print(row_count)
    print(col_count)


    first_str='//*[@id="dataTable"]/tbody/tr['
    second_str=']/td['
    third_str=']'
    i=3


    d={}

    while i<=row_count:
        j=1
        while j<=col_count:
            final_str=first_str+str(i)+second_str+str(j)+third_str
            #Store in dictionary
            d[columns[j-1]]=driver.find_element_by_xpath(final_str).text
            j=j+1
        d[columns[j-1]]=cbd
        j=j+1
        d[columns[j-1]]="Jan24,2018 "
        i=i+1
        print (d)
        df=df.append(d,ignore_index=True)
    f = open(cbd + ".xlsx")
    path = os.path.realpath(f.name)
    book = load_workbook(path)
    writer = pd.ExcelWriter(path, engine='openpyxl')
    writer.book = book
    df.to_excel(writer,sheet_name='NSE UNDERLYING',columns=["Symbol",cbd+" OI","Jan24 OI","Change in OI","% Change in OI","Volume in Contracts","Futures in Crores","Options(Notional) in Crores","Total in Crores","Options(Premium) in Crores","Underlyning Value","Current Business Date","Previous Business Date"])
    writer.save()
    writer.close()
    driver.close()
