from selenium import webdriver
from openpyxl import Workbook
import json
import sys, os, time, traceback


def wait_loading(driver, sec):
    while sec > 0:
        try:
            driver.find_element_by_xpath('//div[@class="footerVisitOtherSites"]')
            return
        except:
            sec -= 1
            time.sleep(1)

def clear_loading(driver):
    try:
        driver.get('blank.html')
    except:
        pass

def capture_screen(driver, name):
    currentdir = os.path.dirname(os.path.realpath(__file__)) + os.sep
    try:
        os.mkdir(currentdir + 'pics')
    except:
        pass
    driver.save_screenshot(currentdir + 'pics' + os.sep + name + '.png')


def validateParam(param):
    try:
        ls = param.split('/')
        if len(ls) != 3: return False
        mon, day, yr = int(ls[0]), int(ls[1]), int(ls[2])
        if mon < 1 or mon > 12 or day < 1 or day > 31 or yr < 2016:
            return False
        return True 
    except:
        return False

def find_pay_button(driver, res):
    cartype = ['Intermediate','Economy','Standard']
    idx = 0
    while idx < len(cartype):
        try:
            locator = '//li[@class="carView"]/div[contains(@class,"brandName")]/h2[text()="' + cartype[idx] + '"]/../..//a[@id="payNowButton"]'
            btn = driver.find_element_by_xpath(locator)
            res['carType'] = cartype[idx]
            return btn
        except:
            idx += 1
    return None

def get_sur_tax_header(driver):
    try:
        locator = '//div[@id="estimationpanel"]//div[@id="taxHeading"]/a'
        return driver.find_element_by_xpath(locator)
    except:
        pass

def base_rate(driver, res):
    cnt = 3
    while cnt > 0:
        try:
            locator = '//div[@id="estimationpanel"]//span[@id="baseRateamountHeading"]/strong'
            val = driver.find_element_by_xpath(locator).text
            res['base_rate'] = val
            return val
        except:
            cnt -= 1

def total_tax_surcharge(driver, res):
    cnt = 3
    while cnt > 0:
        try:
            locator = '//div[@id="estimationpanel"]//span[@id="tx_sur_tot"]/strong'
            val = driver.find_element_by_xpath(locator).text
            res['total_tax_surcharge'] = val
            return val
        except:
            cnt -= 1

def concession_recovery_fee(driver, res):
    cnt = 3
    while cnt > 0:
        try:
            locator = '//div[@id="estimationpanel"]//span[contains(text(),"Concession Recovery Fee")]/following-sibling::span'
            val = driver.find_element_by_xpath(locator).text
            res['concession_recovery_fee'] = val
            return val
        except:
            cnt -= 1

def customer_facility_charge(driver, res):
    cnt = 3
    while cnt > 0:
        try:
            locator = '//div[@id="estimationpanel"]//span[contains(text(),"Customer Facility Charge")]/following-sibling::span[last()]'
            val = driver.find_element_by_xpath(locator).text
            res['customer_facility_charge'] = val
            return val
        except:
            cnt -= 1

def tourism_assessment_fee(driver, res):
    cnt = 3
    while cnt > 0:
        try:
            locator = '//div[@id="estimationpanel"]//span[contains(text(),"Tourism Assessment")]/following-sibling::span'
            val = driver.find_element_by_xpath(locator).text
            res['tourism_assessment_fee'] = val
            return val
        except:
            cnt -= 1

def vehicle_license_fee(driver, res):
    cnt = 3
    while cnt > 0:
        try:
            locator = '//div[@id="estimationpanel"]//span[contains(text(),"Vehicle License Fee")]/following-sibling::span[last()]'
            val = driver.find_element_by_xpath(locator).text
            res['vehicle_license_fee'] = val
            return val
        except:
            cnt -= 1

def total_tax(driver, res):
    cnt = 3
    while cnt > 0:
        try:
            locator = '//div[@id="estimationpanel"]//span[@id="tax_tot"]'
            val = driver.find_element_by_xpath(locator).text
            res['total_tax'] = val
            return val
        except:
            cnt -= 1

def estimated_total(driver, res):
    cnt = 3
    while cnt > 0:
        try:
            locator = '//div[@id="estimationpanel"]//span[@id="estimatedTotal"]'
            val = driver.find_element_by_xpath(locator).text
            res['estimated_total'] = val
            return val
        except:
            cnt -= 1

def save(results):
    if len(results) == 0 : return
    wb = Workbook()
    ws = wb.worksheets[0]
    hdr = list(results[0].keys())
    hdr.sort()
    #make header
    ws.append(hdr)
    for r in results:
        l = [ r[i] for i in hdr]
        ws.append(l)
    wb.save(filename = 'results.xlsx')

home_url = 'https://www.avis.com/car-rental/avisHome/home.ac'

usage = '''
Usage:
<program> <start date> <end date>

example:
scraper.py 02/16/2016 02/17/2016
'''

if __name__ == '__main__': 
    if len(sys.argv) < 3:
        print("Not enough parameters!")
        print()
        print(usage)
        sys.exit()

    if not validateParam(sys.argv[1]) or not validateParam(sys.argv[2]):
        print("Illegal parameters!")
        sys.exit()

    f = open('config.json') 
    cfg = json.load(f)
    f.close()

    driver = webdriver.PhantomJS()
    driver.set_window_size(1024,768)
    results = []
    airports = cfg.get('airports')
    print('Find ' + str(len(airports)) + ' airports')
    for idx, item in enumerate(airports):
        try:
            res = {}
            print('Processing no.' + str(idx+1) + ' ' + item + ' ... ')
            res['airport'] = item
            res['start_date'] = sys.argv[1]
            res['end_date'] = sys.argv[2]
            # home url
            driver.get(home_url)
            wait_loading(driver, 10)
            airport = driver.find_element_by_name('resForm.pickUpLocation')
            airport.clear()
            airport.send_keys(item)
            sdate = driver.find_element_by_name('resForm.pickUpDate')
            sdate.clear()
            sdate.send_keys(sys.argv[1])
            edate = driver.find_element_by_name('resForm.dropOffDate')
            edate.clear()
            edate.send_keys(sys.argv[2])
            queryBtn = driver.find_element_by_id('selectMyCarId')
            queryBtn.click()
            # car url
            time.sleep(1)
            wait_loading(driver, 10)
            # find cars, try Intermediate first, then Economy and Standard
            payBtn = find_pay_button(driver, res)  
            payBtn.click()
            # grab the result
            time.sleep(1)
            wait_loading(driver, 10)
            base_rate(driver,res)
            total_tax_surcharge(driver, res)
            concession_recovery_fee(driver,res)
            customer_facility_charge(driver, res)
            tourism_assessment_fee(driver, res)
            vehicle_license_fee(driver, res)
            total_tax(driver, res)
            estimated_total(driver, res)
            hdr = get_sur_tax_header(driver)
            hdr.click()
            capture_screen(driver, item)
            results.append(res)
            clear_loading(driver)
            print('done!')
        except:
            traceback.print_exc()
            capture_screen(driver,'debug')
            driver.quit()
            sys.exit()

    driver.quit()
    save(results)
    print()
    print('Saving ...', end='')
    print('done!')

