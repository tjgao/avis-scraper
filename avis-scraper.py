from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import os, sys, json, time, copy
from openpyxl import Workbook

def validate_param(param):
    try:
        ls = param.split('/')
        if len(ls) != 3: return False
        mon, day, yr = int(ls[0]), int(ls[1]), int(ls[2])
        if mon < 1 or mon > 12 or day < 1 or day > 31 or yr < 2017:
            return False
        return True 
    except:
        return False

def validate_params(arg1, arg2):
    if not validate_param(arg1) or not validate_param(arg2):
        return False
    return arg1 < arg2


def set_pick_drop(css_selector, browser):
    elem = browser.find_element_by_css_selector(css_selector + '_value')
    elem.send_keys(item)
    wait = WebDriverWait(browser, 10)
    el_addr = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, css_selector + '_dropdown > div.angucomplete-results > div:nth-child(1) > div.angucomplete-row > div')))
    el_addr.click()

def set_pick_drop_date(start, end, browser):
    elem = browser.find_element_by_css_selector('#from')
    elem.send_keys(start)
    elem = browser.find_element_by_css_selector('#to')
    elem.send_keys(end)

def go_to_car_page(browser):
    elem = browser.find_element_by_css_selector('#res-home-select-car')
    elem.click()
    wait = WebDriverWait(browser, 10)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#vehicleTeaser > div.reservation-progress-bar > div > nav > div > ul > li.active')))

def get_car_info(car_types, browser):
    # get all car types
    elems = browser.find_elements_by_xpath('//h3[contains(@ng-bind, "car.carGroup")]')
    elem_types = [e.text for e in elems]
    picked_type = None
    for t in car_types:
        if t not in elem_types: continue
        picked_type = t
        break
    if not picked_type:
        return None
    # find the 'Pay Now' button
    xpath_selector_tpl = '//div//div//div//h3[contains(@ng-bind,"car.carGroup") and contains(text(), "{0}")]//..//..//..//..//div[contains(@class, "paynow")]//a[@id="res-vehicles-pay-now"]'
    xpath_selector = xpath_selector_tpl.format(t)
    btn  = browser.find_element_by_xpath(xpath_selector)
    browser.execute_script("arguments[0].click();", btn)
    wait = WebDriverWait(browser, 10)
    wait.until(EC.visibility_of_element_located((By.ID, 'amt-prepaid')))
    return picked_type

def collector_fee_info(browser):
    results = {}
    elem = browser.find_element_by_xpath('//div//span[contains(text(), "Fees & Taxes")]/following-sibling::span[1]//span[2]')
    results['Total Fees & Taxes'] = elem.text
    elem = browser.find_element_by_xpath('//div//span[contains(text(), "Fees & Taxes")]/following-sibling::div[1]//div//div//span[contains(text(), "Concession Recovery Fee")]/following-sibling::span[1]')
    results['Concession Recovery Fee'] = browser.execute_script('return arguments[0].textContent', elem)
    elem = browser.find_element_by_xpath('//div//span[contains(text(), "Fees & Taxes")]/following-sibling::div[1]//div//div//span[contains(text(), "Customer Facility Charge")]/following-sibling::span[1]')
    results['Customer Facility Charge'] = browser.execute_script('return arguments[0].textContent', elem)
    elem = browser.find_element_by_xpath('//div//span[contains(text(), "Fees & Taxes")]/following-sibling::div[1]//div//div//span[contains(text(), "Energy Recovery Fee")]/following-sibling::span[1]')
    results['Energy Recovery Fee'] = browser.execute_script('return arguments[0].textContent', elem)
    elem = browser.find_element_by_xpath('//div//span[contains(text(), "Fees & Taxes")]/following-sibling::div[1]//div//div//span[contains(text(), "Vehicle License Fee")]/following-sibling::span[1]')
    results['Vehicle License Fee'] = browser.execute_script('return arguments[0].textContent', elem)
    elem = browser.find_element_by_xpath('//div//span[contains(text(), "Fees & Taxes")]/following-sibling::div[1]//div//div//span[contains(text(), "Total Tax")]/following-sibling::span[1]')
    results['Total Tax'] = browser.execute_script('return arguments[0].textContent', elem)
    elem = browser.find_element_by_xpath('//span[contains(@class, "est-total")]//span[2]')
    results['Estimated Total'] = elem.text
    return results

def save(results):
    if len(results) == 0 : return
    excel_hdr = ['Airport', 'Car Type', 'Estimated Total', 'Total Fees & Taxes', 'Total Tax', 'Concession Recovery Fee', 'Customer Facility Charge', 'Energy Recovery Fee', 'Vehicle License Fee']
    wb = Workbook()
    ws = wb.worksheets[0]
    ws.append(excel_hdr)
    for r in results:
        l = [ r[i] for i in excel_hdr]
        ws.append(l)
    wb.save(filename = 'results.xlsx')

usage = '''
Usage:
<program> <start date> <end date>

start date < end date
example:
avis-scraper.py 05/20/2017 05/21/2017
'''

if __name__ == '__main__':
    driver_dir = os.path.dirname(os.path.abspath(__file__))
    if sys.platform == 'win32':	
        driver_path = driver_dir + os.sep + 'chromedriver.exe'
    elif sys.platform == 'darwin':
        driver_path = driver_dir + os.sep + 'chromedriver'
    else:
        print('Unsupported platform!')
        sys.exit()

    if len(sys.argv) < 3:
        print(usage)
        sys.exit()

    if not validate_params(sys.argv[1], sys.argv[2]):
        print("Illegal parameters!")
        print(usage)
        sys.exit()		

    with open('config.json') as f:
        airports = json.load(f).get('airports')
    # start browser
    browser = webdriver.Chrome(driver_path)

    car_types = ['Economy', 'Compact', 'Intermediate', 'Standard']
    results = []
    for idx, item in enumerate(airports):
        browser.get('https://www.avis.com/en/home')
        print('Processing No.' + str(idx + 1) + ' ' + item + ' ...   ', end='')	
        set_pick_drop('#PicLoc', browser)
        set_pick_drop('#DropLoc', browser)
        set_pick_drop_date(sys.argv[1], sys.argv[2], browser)
        go_to_car_page(browser)
        cartype = get_car_info(car_types, browser)
        if not cartype:
            print('Failed: Could not find specified types')
            continue
        car_info = collector_fee_info(browser)
        car_info['Car Type'] = cartype
        car_info['Airport'] = item
        results.append(car_info)
        print('done!')
    browser.quit()
    save(results)

