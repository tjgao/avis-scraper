# Note: this code does not work any more since avis changed their website
from lxml import html
from openpyxl import Workbook
import json
import sys, traceback
import requests

DEBUG = 0

home_url = 'https://www.avis.com/car-rental/avisHome/home.ac'
query_url = 'https://www.avis.com/car-rental/reservation/time-place-submit.ac'
pay_url = 'https://www.avis.com/car-rental/reservation/select-car.ac'

excel_hdr = ['airport','carType','start_date','end_date','base_rate','total_tax_surcharge','concession_recovery_fee','concession_recovery_fee_surcharge','tourism_assessment_fee','customer_facility_charge','energy_recovery_fee','vehicle_license_fee','transportation_fee','total_tax','estimated_total']

queryfm = {
"resForm.pickUpLocation":"LAX",
"pickupKeywordValue":None,
"pickupSuggestionValue":None,
"pickupCityFlag":"false",
"pickupSelectedValue":None,
"__checkbox_resForm.returnSameAsPickUp.value":"true",
"resForm.dropOffLocation":"Airport, city, zip, address, attraction",
"hiddenRtnLoc":None,
"dropoffKeywordValue":None,
"dropoffSuggestionValue":None,
"dropoffCityFlag":"false",
"dropoffSelectedValue":None,
"resForm.pickUpDate":None,
"resForm.pickUpTime":"12:00 PM",
"resForm.dropOffDate":None,
"resForm.dropOffTime":"12:00 PM",
"resForm.age":"25+",
"resForm.USResident.value":"true",
"__checkbox_resForm.USResident.value":"true",
"resForm.countryOfRes.value":"US",
"__checkbox_resForm.wizardInfo.value":"true",
"resForm.wizardNumber.value":None,
"resForm.lastName.value":None,
"__checkbox_resForm.promotionInfo.value":"true",
"resForm.awdRateCode.value":None,
"resForm.awdNonUSRateCode.value":None,
"resForm.couponCode":None,
"resForm.rateCode.value":None,
"resForm.reservationDisplayVO.userEnteredCoupon":"false",
"resForm.clearPromotionInfo":"true",
"resForm.paperCoupon":"false",
"resForm.samePage":"avisHome",
"resForm.vehicleCategoryType":None,
"resForm.vehicleCategoryDescription":None,
"resForm.vehicleCountry":None,
"resForm.truckSuggestion":"false"
}

payfm = {
"resForm.signatureCarFeatured":"true",
"isChatBoxClicked":"false",
"resForm.rateCodeEnforced":"false",
"resForm.carGroupCode":None,
"resForm.prepayDiscount":None,
"resForm.soldOutCarGroupCode":None,
"resForm.userSelectedRateCode":None,
"resForm.prepay":"true",
"resForm.pickUpLocation":"Los Angeles Intl Airport-(LAX)",
"resForm.dropOffLocation":"Los Angeles Intl Airport-(LAX)",
"resForm.returnSameAsPickUp.value":"false",
"originalcurrency":"USD",
"changecurrency":"USD",
"resForm.carSelected":None,
"resForm.reservationDisplayVO.userEnteredCoupon":"false",
"isChatBoxClicked":"false",
"resForm.reservationDisplayVO.upsellCar.upsellEligibleOnStepThree":"false"    
}

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
 

def find_btn_val(doc, res):
    cartype = ['Intermediate','Economy','Standard']
    res['carType'] = ''
    idx = 0
    try:
        # make sure we are on the right page
        locator = '//li[@class="carView"]'
        e = doc.xpath(locator)[0]
    except:
        return None
    # try 'pay now'
    while idx < len(cartype):
        try:
            payfm['resForm.signatureCarFeatured'] = 'true'
            locator = '//li[@class="carView"]/div[contains(@class,"brandName")]/h2[text()="' + cartype[idx] + '"]/../..//a[@id="payNowButton"]/@onclick'
            e = doc.xpath(locator)[0]
            res['carType'] = cartype[idx]
            ee = e[ e.find('(') + 1 : e.find(')')]
            ls = [ i.strip().strip("'") for i in ee.split(',') ]
            return ls
        except:
            #traceback.print_exc()
            #print('paynow fail')
            idx += 1
    # try 'pay later'
    idx = 0
    while idx < len(cartype):
        try:
            payfm['resForm.signatureCarFeatured'] = 'false'
            locator = '//li[@class="carView"]/div[contains(@class,"brandName")]/h2[text()="' + cartype[idx] + '"]/../..//a[@id="selectPayLaterDom"]/@onclick'
            e = doc.xpath(locator)[0]
            res['carType'] = cartype[idx]
            ee = e[ e.find('(') + 1 : e.find(')')]
            ls = [ i.strip().strip("'") for i in ee.split(',') ]
            return ls
        except:
            #traceback.print_exc()
            #print('paylater fail')
            idx += 1
    return None

# important parameters hidden in javascript 
# for example: javascript:submitForm('C','LC','true', '7.400000000000006' );
def prepare_form(doc, res):
    try:
        ls = find_btn_val(doc, res)
        if ls is None: return False
        locator = '//span[@class="locDetails"]'
        e = doc.xpath(locator)
        try:
            payfm['resForm.pickUpLocation'] = e[0].text.strip()
            payfm['resForm.dropOffLocation'] = e[1].text.strip()
            payfm['resForm.carGroupCode'] = ls[0]
            payfm['resForm.userSelectedRateCode'] = ls[1]
            payfm['resForm.prepay'] = ls[2]
            if ls[2] == 'true': payfm['resForm.prepayDiscount'] = ls[3]
            else: payfm['resForm.prepayDiscount'] = ''
        except:
            return True
        return True
    except:
        #traceback.print_exc()
        print('pay button not found')
    return False


def dPrint(s):
    if DEBUG == 0: return
    if DEBUG == 1: print(s)
    else: traceback.print_exc()


def base_rate(doc, res):
    try:
        res['base_rate'] = 0
        locator = '//div[@id="estimationpanel"]//span[@id="baseRateamountHeading"]/strong'
        val = doc.xpath(locator)[0].text.strip()
        res['base_rate'] = round(float(val),2)
        return val
    except:
        #traceback.print_exc()
        dPrint('Fail to get base rate')

def total_tax_surcharge(doc, res):
    try:
        res['total_tax_surcharge'] = 0
        locator = '//div[@id="estimationpanel"]//span[@id="tx_sur_tot"]/strong'
        val = doc.xpath(locator)[0].text.strip()
        res['total_tax_surcharge'] = round(float(val),2)
        return val
    except:
        #traceback.print_exc()
        dPrint('Fail to get total tax & surcharge ')

def concession_recovery_fee(doc, res):
    try:
        res['concession_recovery_fee'] = 0
        locator = '//div[@id="estimationpanel"]//span[contains(text(),"Concession Recovery Fee")]/following-sibling::span'
        val = doc.xpath(locator)[0].text.strip()
        res['concession_recovery_fee'] = round(float(val),2)
        return val
    except:
        #traceback.print_exc()
        dPrint('Fail to get concession recovery fee')

def concession_recovery_fee_surcharge(doc, res):
    try:
        res['concession_recovery_fee_surcharge'] = 0
        locator = '//div[@id="estimationpanel"]//span[contains(text(),"Concession Recovery Fee Surcharge")]/following-sibling::span[last()]'
        val = doc.xpath(locator)[0].text.strip()
        res['concession_recovery_fee_surcharge'] = round(float(val),2)
    except:
        dPrint('Fail to get concession recovery fee surcharge')

def customer_facility_charge(doc, res):
    try:
        res['customer_facility_charge'] = 0
        locator = '//div[@id="estimationpanel"]//span[contains(text(),"Customer Facility Charge")]/following-sibling::span[last()]'
        val = doc.xpath(locator)[0].text.strip()
        res['customer_facility_charge'] = round(float(val),2)
        return val
    except:
        #traceback.print_exc()
        dPrint('Fail to get customer facility charge')

def tourism_assessment_fee(doc, res):
    try:
        res['tourism_assessment_fee'] = 0
        locator = '//div[@id="estimationpanel"]//span[contains(text(),"Tourism Assessment")]/following-sibling::span[last()]'
        val = doc.xpath(locator)[0].text.strip()
        res['tourism_assessment_fee'] = round(float(val),2)
        return val
    except:
        #traceback.print_exc()
        dPrint('Fail to get tourism assessment fee')

def transportation_fee(doc, res):
    try:
        res['transportation_fee'] = 0
        locator = '//div[@id="estimationpanel"]//span[contains(text(),"Transportation Fee")]/following-sibling::span[last()]'
        val = doc.xpath(locator)[0].text.strip()
        res['transportation_fee'] = round(float(val),2)
        return val
    except:
        dPrint('Fail to get transportation fee')

def energy_recovery_fee(doc, res):
    try:
        res['energy_recovery_fee'] = 0
        locator = '//div[@id="estimationpanel"]//span[contains(text(),"Energy Recovery Fee")]/following-sibling::span[last()]'
        val = doc.xpath(locator)[0].text.strip()
        res['energy_recovery_fee'] = round(float(val), 2)
        return val
    except:
        dPrint('Fail to get energy recovery fee')
def vehicle_license_fee(doc, res):
    try:
        res['vehicle_license_fee'] = 0
        locator = '//div[@id="estimationpanel"]//span[contains(text(),"Vehicle License Fee")]/following-sibling::span[last()]'
        val = doc.xpath(locator)[0].text.strip()
        res['vehicle_license_fee'] = round(float(val),2)
        return val
    except:
        #traceback.print_exc()
        dPrint('Fail to get vehicle license fee')

def total_tax(doc, res):
    try:
        res['total_tax'] = 0
        locator = '//div[@id="estimationpanel"]//span[@id="tax_tot"]'
        val = doc.xpath(locator)[0].text.strip()
        res['total_tax'] = round(float(val),2)
        return val
    except:
        #traceback.print_exc()
        dPrint('Fail to get total tax')

def estimated_total(doc, res):
    try:
        res['estimated_total'] = 0
        locator = '//div[@id="estimationpanel"]//span[@id="estimatedTotal"]'
        val = doc.xpath(locator)[0].text.strip()
        res['estimated_total'] = round(float(val),2)
        return val
    except:
        #traceback.print_exc()
        dPrint('Fail to get estimated total')

def save(results):
    if len(results) == 0 : return
    wb = Workbook()
    ws = wb.worksheets[0]
    #t = list(results[0].keys())
    #make header
    ws.append(excel_hdr)
    for r in results:
        l = [ r[i] for i in excel_hdr]
        ws.append(l)
    wb.save(filename = 'results.xlsx')

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
    airports = cfg.get('airports')

    results = []
    session = requests.Session()
    session.get(home_url)
    print(str(len(airports)) + ' airports found')
    for idx, item in enumerate(airports):
        res = {}
        print('Processing no.' + str(idx+1) + ' ' + item + ' ... ')
        res['airport'] = item
        res['start_date'] = sys.argv[1]
        res['end_date'] = sys.argv[2]        
        queryfm['resForm.pickUpLocation'] = item
        queryfm['resForm.pickUpDate'] = sys.argv[1]
        queryfm['resForm.dropOffDate'] = sys.argv[2]
        resp = session.post(query_url, data = queryfm)
        doc = html.fromstring(resp.text)
        if not prepare_form(doc, res):
            print('Fail to find "Intermediate", "Economy" or "Standard" type for ' + item)
            continue
        resp = session.post(pay_url, data = payfm )
        doc = html.fromstring(resp.text)
        #with open('out.txt','w') as f:
        #    f.write(resp.text)
        base_rate(doc, res)
        total_tax_surcharge(doc, res)
        concession_recovery_fee(doc, res)
        concession_recovery_fee_surcharge(doc, res)
        customer_facility_charge(doc, res)
        tourism_assessment_fee(doc, res)
        energy_recovery_fee(doc, res)
        vehicle_license_fee(doc, res)
        transportation_fee(doc,res)
        total_tax(doc, res)
        estimated_total(doc, res)
        #print(res)
        results.append(res)
        print('Done!')

    print()
    print('Saving results ... ')
    save(results)
    print('Done!')
