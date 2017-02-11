#!/bin/python

import json, requests
import sys, traceback
from openpyxl import Workbook

home_url = 'https://www.avis.com/en/home'
car_reservation_url = 'https://www.avis.com/webapi/v1/reservation/vehicles'
car_extras_url = 'https://www.avis.com/webapi/v1/reservation/extras'


car_reservation_req = {
   "rqHeader":{
	  "brand":"",
	  "locale":""
   },
   "pickInfo":"",  # need to fill
   "pickDate":"",  # need to fill
   "pickTime":"12:00 PM",
   "dropInfo":"",  # need to fill
   "dropDate":"",  # need to fill
   "dropTime":"12:00 PM",
   "couponNumber":"",
   "couponInstances":"",
   "discountNumber":"",
   "rateType":"",
   "residency":"US",
   "age":25,
   "wizardNumber":"",
   "lastName":"",
   "userSelectedCurrency":""
}

car_extras_req = {
   "rqHeader":{
	  "locale":"en_US"
   },
   "carClass":"", # need to fill
   "rateCode":"", # need to fill
   "prepayProcessed":"true",
   "coupon":"",
   "discountNumber":"",
   "couponInstances":"",
   "selectedInsurances":[],
   "selectedProducts":[],
   "deSelectedProducts":[],
   "selectedNonSupCPlist":[],
   "updateUserProducts":"false",
   "enforceRateCode":"false"
}

usage = '''
Usage:
<program> <start date> <end date>

start date < end date
example:
scraper.py 02/16/2016 02/17/2016
'''
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

def generate_file_name(arg1, arg2):
	arg1.replace('/','-')
	arg2.replace('/','-')
	return 'results_' + arg1 + '_' + arg2 + '.xlsx'

# intermediate -> economy -> standard -> full size
# if still no match, then just pick the first one
def find_carclass_ratecode(j):
	vehicle_list = j['vehicleSummaryList']
	car_types = {'Intermediate':None, 'Economy':None, 'Standard':None, 'Full Size':None, 'Other':None}
	for o in vehicle_list:
		u = o['carGroup'] 
		if u == 'Intermediate' and o['carAvailability'] == 'A':
			return (o['carClass'], o['rateCode'], o['carGroup'])
		elif u in car_types and o['carAvailability'] == 'A':
			car_types[u] = (o['carClass'], o['rateCode'], o['carGroup'])
		elif o['carAvailability'] == 'A' and car_types['Other'] is None:
			car_types['Other'] = (o['carClass'], o['rateCode'], o['carGroup'])

	if car_types['Economy']: return car_types['Economy']
	if car_types['Standard']: return car_types['Standard']
	if car_types['Full Size']: return car_types['Full Size']
	if car_types['Other']: return car_types['Other']
	return None	

excel_header = ['airport', 'car type', 'estimatedTotal', 
'baseRate', 'surchargeTotal', 'totalTax', 'Concession Recovery Fe',
'Customer Facility Charge', 'Tourism Assessment Fee', 'Vehicle License Fee'
]

def fill_record(record,j):
	s = j['rateSummary']
	record['estimatedTotal'] = s['estimatedTotal']
	record['baseRate'] = s['baseRate']
	record['surchargeTotal'] = s['surchargeTotal']
	record['totalTax'] = s['totalTax']
	# Make sure the order is correct
	s['surcharges'].sort(key=lambda x:x['name'])
	record['Concession Recovery Fee'] = s['surcharges'][0]
	record['Customer Facility Charge'] = s['surcharges'][1]
	record['Tourism Assessment Fee'] = s['surcharges'][2]
	record['Vehicle License Fee'] = s['surcharges'][3]

def save_to_disk(results, arg1, arg2):
	if len(results) == 0:
		print('No data to save')
		return
	wb = Workbook()
	ws = wb.worksheets[0]
	ws.append(excel_header)
	for r in results:
		l = [r[i] for i in excel_header]
		ws.append(l)
	wb.save(filename = generate_file_name(arg1, arg2))

if __name__== '__main__':
	if len(sys.argv) < 3:
		print(usage)
		sys.exit()
	
	if not validate_params(sys.argv[1], sys.argv[2]):
		print("Illegal parameters!")
		print(usage)
		sys.exit()

	with open('config.json') as f:
		airports = json.load(f).get('airports')

	print(str(len(airports)) + ' airports found')
	results = []
	session = requests.Session()
	print(session.cookies.get_dict())
	resp = session.get(home_url)
	for idx, item in enumerate(airports):
		print('Processing No.' + str(idx + 1) + ' ' + item + ' ...   ', end='')
		car_reservation_req['pickInfo'] = item
		car_reservation_req['pickDate'] = sys.argv[1]
		car_reservation_req['dropDate'] = sys.argv[2]
		resp = session.post(car_reservation_url, json = car_reservation_req)
		print(resp.text)
		cars = json.loads(resp.text)
		car_class_ratecode = find_carclass_coderate(cars)
		if car_class_ratecode is None:
			print('')
			print('Could not find any car for ' + item)
			continue
		car_extras_req['carClass'] = car_class_ratecode[0]
		car_extras_req['rateCode'] = car_class_ratecode[1]
		resp = session.post(car_extras_url, json = car_extras_req).text
		car_details = json.loads(resp)
		res = {'airport':item}
		fill_record(res, car_details)
		results.append(res)
		print(' done.')
	print('')
	print('Saving results ... ', end='')
	save_to_disk(results, sys.argv[1], sys.argv[2])
	print(' all done!')