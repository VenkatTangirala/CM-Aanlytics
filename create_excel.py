from selenium import webdriver
from selenium.webdriver.support import ui
from selenium.webdriver.support.ui import WebDriverWait
import xlsxwriter
import time
from . import nse_52_low
from . import BSE_52_HIGH
from . import option_chain,nse_52_high,nse_both_band,nse_contracts,nse_lower_band,nse_underlyning,nse_upper_band,nse_volume
def page_is_loaded(driver):
    # this function waits for the page to load before performing further actions.
    return driver.find_element_by_tag_name("body") != None

def create():
    driver = webdriver.Chrome(r'E:\chromedriver.exe')
    driver.get("https://www.nseindia.com/products/content/equities/equities/oi_spurts.htm")
    time.sleep(4)

    wait = ui.WebDriverWait(driver,10)
    wait.until(page_is_loaded)

    cdate=driver.find_element_by_xpath('//*[@id="time"]')
    arr = cdate.text.split(' ')
    res = arr[2]+arr[3]+arr[4]
    res = str(res)
    print(res)
# Create an new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook(res+'.xlsx')
    worksheet = workbook.add_worksheet()
    workbook.close()
    return res

def func(date):
    #nse_52_low.nse_low(date)
    #BSE_52_HIGH.bse_high(date)
    option_chain.option(date)
    #nse_contracts.contracts(date)
    #nse_underlyning.underlying(date)
    #nse_volume.volume(date)
    #nse_52_high.nse_52_hi(date)
    #nse_upper_band.upper_band(date)
    #nse_lower_band.lower_band(date)
    #nse_both_band.both_bands(date)

#date=create()
#func(date)