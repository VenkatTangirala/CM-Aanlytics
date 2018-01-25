from selenium import webdriver
from selenium.webdriver.support import ui
from selenium.webdriver.support.ui import WebDriverWait
from pandas import ExcelWriter
import pandas as pd
from openpyxl import load_workbook
import datetime
import time
import os


def page_is_loaded(driver):
    # this function waits for the page to load before performing further actions.
    return driver.find_element_by_tag_name("body") != None


def lower_band(ris):
    print("NSE LOWER BAND")
    driver = webdriver.Chrome(r'E:\chromedriver.exe')

#driver.get("https://www.zaubacorp.com/")
    driver.get("https://www.nseindia.com/products/content/equities/equities/price_band_hitters.htm")
    columns=["Symbol","Series","LTP","Change","%Change","Price Band%","High","Low","52Week High","52Week Low","Volume(Shares)","Value(in Lacs)","View","Current Business Date"]
    df=pd.DataFrame()
    time.sleep(5)
    button=driver.find_element_by_xpath('//*[@id="tab8"]')
    button.click()

    time.sleep(5)
    row_count=len(driver.find_elements_by_xpath('//*[@id="dataTable"]/tbody/tr'))
    col_count=len(driver.find_elements_by_xpath('//*[@id="dataTable"]/tbody/tr[2]/td'))

    print (row_count)
    print (col_count)
    #f = open(ris+".xlsx")
    #st=os.path.realpath(f.name)

    first_str='//*[@id="dataTable"]/tbody/tr['
    second_str=']/td['
    third_str=']'

    i=2
    d={}
    while i<=row_count:
        j=1
        while j<=col_count:
            final_str=first_str+str(i)+second_str+str(j)+third_str
            d[columns[j-1]]=driver.find_element_by_xpath(final_str).text
            j=j+1
        d[columns[j-1]]='All'
        j=j+1
        d[columns[j-1]]=ris
        df=df.append(d,ignore_index=True)
        i=i+1

    button=driver.find_element_by_xpath('//*[@id="G"]')
    button.click()
    time.sleep(5)
    row_count=len(driver.find_elements_by_xpath('//*[@id="dataTable"]/tbody/tr'))
    col_count=len(driver.find_elements_by_xpath('//*[@id="dataTable"]/tbody/tr[2]/td'))



    print (row_count)
    print (col_count)
    i=2
    while i<=row_count:
        j=1
        while j<=col_count:
            final_str=first_str+str(i)+second_str+str(j)+third_str
            d[columns[j-1]]=driver.find_element_by_xpath(final_str).text
            j=j+1
        d[columns[j-1]]='Securities > Rs. 20'
        j=j+1
        d[columns[j-1]] = ris
        df=df.append(d,ignore_index=True)
        i=i+1

    button=driver.find_element_by_xpath('//*[@id="L"]')
    button.click()

    time.sleep(5)
    row_count=len(driver.find_elements_by_xpath('//*[@id="dataTable"]/tbody/tr'))
    col_count=len(driver.find_elements_by_xpath('//*[@id="dataTable"]/tbody/tr[2]/td'))

    i=2
    while i<=row_count:
        j=1
        while j<=col_count:
            final_str=first_str+str(i)+second_str+str(j)+third_str
            d[columns[j-1]]=driver.find_element_by_xpath(final_str).text
            j=j+1
        d[columns[j-1]]='Securities < Rs. 20'
        j=j+1
        d[columns[j-1]]=ris
        df=df.append(d,ignore_index=True)
        i=i+1

    #path=st
    #book = load_workbook(path)
    #writer = pd.ExcelWriter(path, engine = 'openpyxl')
    #writer.book = book
    #writer=pd.ExcelWriter(tdate+".xlsx")
    df.to_excel("lowerband.xls",columns=["Symbol","Series","LTP","Change","%Change","Price Band%","High","Low","52Week High","52Week Low","Volume(Shares)","Value(in Lacs)","View","Current Business Date"])
    #writer.save()
    #writer.close()
    #df1.to_excel("nse_india_name_low.xls")
    driver.close()
