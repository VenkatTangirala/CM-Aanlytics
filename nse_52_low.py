from selenium import webdriver
from selenium.webdriver.support import ui
from selenium.webdriver.support.ui import WebDriverWait
import pandas as pd
from openpyxl import load_workbook
import time
import os
def page_is_loaded(driver):
    # this function waits for the page to load before performing further actions.
    return driver.find_element_by_tag_name("body") != None


def nse_low(date):
    print ("NSE 52 LOW")
    driver = webdriver.Chrome(r'E:\chromedriver.exe')
    driver.get("https://www.nseindia.com/products/content/equities/equities/eq_new_high_low.htm")

    time.sleep(4)

    wait = ui.WebDriverWait(driver, 10)
    wait.until(page_is_loaded)

    ##END OF UNDERLYNING
    button=driver.find_element_by_xpath('//*[@id="tab8"]')
    button.click()
    time.sleep(5)
    df1 = pd.DataFrame()
    columns1 = ["Symbol","Security Name","New-52L","Previous-Low","Previous-LowDate","LTP","Previousclose","change","%change","Current Business Date"]
    #tdate=driver.find_element_by_xpath('//*[@id="time52Week"]')
    #tdate.click()
    #arr=tdate.text.split(' ')
    #ris=arr[2]+arr[3]+arr[4]
    #ris=str(ris)

    f = open(date+".xlsx")
    path = os.path.realpath(f.name)
    #tdate=datetime.date.today()
    #tdate=str(tdate)
    #f = open(res+".xlsx")
    #st=os.path.realpath(f.name)

    k=0
    d1={}
    first_str='//*[@id="replacetext"]/table/tbody/tr['
    second_str=']/td['
    third_str=']'
    print("Started")
    s=0
    num=driver.find_element_by_xpath('//*[@id="pageText"]')
    arr = num.text.split(' ')
    res=int(arr[9])
    while s<res:
        row_count=len(driver.find_elements_by_xpath('//*[@id="replacetext"]/table/tbody/tr'))
        col_count=len(driver.find_elements_by_xpath('//*[@id="replacetext"]/table/tbody/tr[3]/td'))
        i=3
        while i<=row_count:
            j=1
            while j<=col_count:
                final_str=first_str+str(i)+second_str+str(j)+third_str
                d1[columns1[j-1]]=driver.find_element_by_xpath(final_str).text
                j=j+1
            i=i+1
            d1[columns1[j-1]]=date
            #print("In Here")
            print (d1)
            df1=df1.append(d1, ignore_index=True)
            button=driver.find_element_by_xpath('//*[@id="pageText"]/a[2]')
            button.click()
            time.sleep(5)
            s=s+1

    book = load_workbook(path)
    writer = pd.ExcelWriter(path, engine = 'openpyxl')
    writer.book = book
    #writer=pd.ExcelWriter(tdate+".xlsx")
    df1.to_excel(writer,sheet_name='NSELOW',columns=["Symbol","Security Name","New-52L","Previous-Low","Previous-LowDate","LTP","Previousclose","change","%change","Current Business Date"])
    writer.save()
    writer.close()
    #df1.to_excel("nse_india_name_low.xls")
    driver.close()
