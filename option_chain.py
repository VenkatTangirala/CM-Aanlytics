from selenium import webdriver
from selenium.webdriver.support import ui
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from openpyxl import load_workbook
import pandas as pd
import os
import datetime 

import time


def page_is_loaded(driver):
    # this function waits for the page to load before performing further actions.
    return driver.find_element_by_tag_name("body") != None

def option(bus_date):
    print ("OPTIN CHAIN")
    driver = webdriver.Chrome(r'E:\chromedriver.exe')

    columns=["CBD OI","Change in OI","Volume","IV","LTP","Net Change","Bid Quantity","Bid Price","Ask Price","Ask Qunatity","Strike Price","Current Business Date","Category","Expiry Date","Stock Market Index"]
    df=pd.DataFrame()


    driver.get('https://www.nseindia.com/live_market/dynaContent/live_watch/option_chain/optionKeys.jsp?symbol=NIFTY&instrument=-&date=-')

    time.sleep(6)

    stock_mark_ind=[]

    select=Select(driver.find_element_by_id('optnContract'))
#select.select_by_value('NIFTY')
    opt=select.options
    for i in opt:
        stock_mark_ind.append((i.text))

    stock_mark_ind=stock_mark_ind[1:]
    print (stock_mark_ind)

    print (driver.find_element_by_xpath('//*[@id="octable"]/tbody/tr[1]/td[13]').text)
    #print (row_count)
    #print (col_count)
    first_str='//*[@id="octable"]/tbody/tr['
    second_str=']/td['
    third_str=']'

    d={}
    wait = ui.WebDriverWait(driver, 10)
    wait.until(page_is_loaded)
    for value in stock_mark_ind:
        print(value)
        i=1
        cols=12
        time.sleep(6)
        print ("after sleep")
        select=Select(driver.find_element_by_id('optnContract'))
        select.select_by_value(value)
        print ("after select")
        wait = ui.WebDriverWait(driver, 10)
        print ("wait")
        wait.until(page_is_loaded)
        print ("after wait")
        time.sleep(10)
        exp_dat=[]
        days=Select(driver.find_element_by_id('date'))
        exp_days=days.options
        for ind in exp_days:
            exp_dat.append((ind.text))
        exp_dat=exp_dat[1:]
        #print (exp_dat)
        print (exp_dat)
        if(len(exp_dat)==0):
            j=2
            while j<=cols:
                    final_str=first_str+str(i)+second_str+str(j)+third_str
                    d[columns[j-2]]="No contracts traded today"
                    j=j+1
            
            d[columns[j-2]]=bus_date
            j=j+1
            d[columns[j-2]]="CALLS"
            j=j+1
            d[columns[j-2]]="No expirey date "
            j=j+1
            d[columns[j-2]]=value
            df=df.append(d,ignore_index=True)
            j=22
            while j>=17:
                    final_str=first_str+str(i)+second_str+str(j)+third_str
                    d[columns[22-j]]="No contracts traded today"
                    j=j-1
            j=13
            while j<=16:
                final_str=first_str+str(i)+second_str+str(j)+third_str
                d[columns[j-7]]="No contracts traded today"
                j=j+1
            d[columns[j-7]]="No contracts traded today"
            j=j+1
            d[columns[j-7]]=bus_date
            j=j+1
            d[columns[j-7]]='PUTS'
            j=j+1
            d[columns[j-7]]="No expirey dates"
            j=j+1
            d[columns[j-7]]=value
            i=i+1
            df=df.append(d,ignore_index=True)
            print (d)
        
        else:
            for expirey in exp_dat:
                i=1
                days=Select(driver.find_element_by_id('date'))
                time.sleep(10)
                days.select_by_value(expirey)
                wait = ui.WebDriverWait(driver, 10)
                wait.until(page_is_loaded)
                print (expirey)
                row_count=len(driver.find_elements_by_xpath('//*[@id="octable"]/tbody/tr'))
                col_count=len(driver.find_elements_by_xpath('//*[@id="octable"]/tbody/tr[1]/td'))
                print (row_count)
                print (col_count)
                while i<row_count:
                    j=2
                    while j<=cols:
                        final_str=first_str+str(i)+second_str+str(j)+third_str
                        d[columns[j-2]]=driver.find_element_by_xpath(final_str).text
                        j=j+1
                
                    d[columns[j-2]]=bus_date
                    j=j+1
                    d[columns[j-2]]="CALLS"
                    j=j+1
                    d[columns[j-2]]=expirey
                    j=j+1
                    d[columns[j-2]]=value
                    df=df.append(d,ignore_index=True)
                    j=22
                    while j>=17:
                        final_str=first_str+str(i)+second_str+str(j)+third_str
                        d[columns[22-j]]=driver.find_element_by_xpath(final_str).text
                        j=j-1
                    j=13
                    while j<=16:
                        final_str=first_str+str(i)+second_str+str(j)+third_str
                        d[columns[j-7]]=driver.find_element_by_xpath(final_str).text
                        j=j+1
                    d[columns[j-7]]=driver.find_element_by_xpath('//*[@id="octable"]/tbody/tr[1]/td[12]/a/b').text
                    j=j+1
                    d[columns[j-7]]=bus_date
                    j=j+1
                    d[columns[j-7]]='PUTS'
                    j=j+1
                    d[columns[j-7]]=expirey
                    j=j+1
                    d[columns[j-7]]=value
                    i=i+1
                    df=df.append(d,ignore_index=True)

    f = open(bus_date + ".xlsx")
    path = os.path.realpath(f.name)
    book = load_workbook(path)
    writer = pd.ExcelWriter(path, engine='openpyxl')
    writer.book = book
    df.to_excel(writer, sheet_name='OPTION CHAIN', columns=["CBD OI","Change in OI","Volume","IV","LTP","Net Change","Bid Quantity","Bid Price","Ask Price","Ask Qunatity","Strike Price","Current Business Date","Category","Expiry Date","Stock Market Index"])
    writer.save()
    writer.close()
    driver.close()


