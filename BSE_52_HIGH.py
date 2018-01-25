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


def bse_high(date):
    print ("BSE 52 HIGH")
    driver = webdriver.Chrome(r'E:\chromedriver.exe')
    columns = ["Security Code", "Security Name", "LTP", "52 Weeks HIGH", "Previous 52 Weeks High(Price/Date)","All Time High(Price/Date)", "Current business Date"]
    df = pd.DataFrame()

    driver.get("http://www.bseindia.com/markets/equity/EQReports/HighLow.aspx")
    time.sleep(4)
    wait = ui.WebDriverWait(driver, 10)
    wait.until(page_is_loaded)

    button = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_lnkHigh"]')
    button.click()

    row_count = len(driver.find_elements_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_GrdvwHigh"]/tbody/tr'))
    col_count = len(driver.find_elements_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_GrdvwHigh"]/tbody/tr[2]/td'))

    print (row_count)
    print (col_count)

    first_str='//*[@id="ctl00_ContentPlaceHolder1_GrdvwHigh"]/tbody/tr['
    second_str = ']/td['
    third_str=']'
    i=2
    d={}
    while i<row_count:
        j=1
        while j<=col_count:
            final_str=first_str+str(i)+second_str+str(j)+third_str
            d[columns[j-1]]=driver.find_element_by_xpath(final_str).text
            j=j+1
        d[columns[j-1]] = date
        df = df.append(d,ignore_index=True)
        i=i+1
    


    print(row_count)
    print(col_count)
    f = open(date + ".xlsx")
    path = os.path.realpath(f.name)
    book = load_workbook(path)
    writer = pd.ExcelWriter(path, engine='openpyxl')
    writer.book = book
    df.to_excel(writer,sheet_name="BSE 52 HIGH", columns=["Security Code","Security Name","LTP","52 Weeks HIGH","Previous 52 Weeks High(Price/Date)","All Time High(Price/Date)","Current business Date"])
    writer.save()
    writer.close()
    driver.close()

#bse_high()
        
