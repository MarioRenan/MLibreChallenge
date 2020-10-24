import json
import sys
import xlwt
import csv
from openpyxl import Workbook

wb = xlwt.Workbook(encoding="UTF-8") #Abrindo Excel
ws = wb.add_sheet('db', cell_overwrite_ok=True) 

ws.write(0, 0, 'dn_name')
ws.write(0, 1, 'owner_uid')
ws.write(0, 2, 'owner_name')
ws.write(0, 3, 'owner_email')
ws.write(0, 4, 'classification')
ws.write(0, 5, 'time_stamp')
ws.write(0, 6, 'manager_mail')

#with open('user_manager.csv') as csv_file:
	#csv = csv.reader(csv_file, delimiter=',')
	#for row in csv:
		#print(row[0]) #row_id
		#print(row[1]) #user_id
		#print(row[2]) #user_state
		#print(row[3]) #user_manager

with open('c:\\Challenge Compliance\\files\\dblist.json') as json_db:
	data = json.load(json_db)

	quantidade = 2

	for count in data['db_list'][0]['dn_name']:
		quantidade +=1

	#print(json.dumps(data, indent=4, sort_keys=True))
	#print(data['db_list'][0]['classification']['confidentiality']) #classificação
	#print(data['db_list'][0]['dn_name']) #dn_name
	#print(data['db_list'][0]['owner']['email']) #email do owner
	#print(data['db_list'][0]['owner']['name']) #name do owner
	#print(data['db_list'][0]['owner']['uid']) #uid do owner
	#print(data['db_list'][0]['time_stamp']) #data e hora


i = 0

while i < quantidade:

	try:
		row = ws.row(i+1)
	except Exception as e:
		print(e)

	try:
		row.write(0, data['db_list'][i]['dn_name'])
	except Exception as e:
		row.write(0, 'N/A')

	try:
		row.write(1, data['db_list'][i]['owner']['uid'])
	except Exception as e:
		row.write(1, 'N/A')

	try:
		row.write(2, data['db_list'][i]['owner']['name'])
	except Exception as e:
		row.write(2, 'N/A')

	try:
		row.write(3, data['db_list'][i]['owner']['email'])
	except Exception as e:
		row.write(3, 'N/A')
		"""
		with open('user_manager.csv', 'rt') as f:
			manager = csv.reader(f, delimiter=',')
			for rows in manager:
				if data['db_list'][i]['owner']['uid'] == rows[1]:
					row.write(3, rows[3])"""

	try:
		row.write(4, data['db_list'][i]['classification']['confidentiality'])
	except Exception as e:
		row.write(4, 'N/A')

	try:
		row.write(5, data['db_list'][i]['time_stamp'])
	except Exception as e:
		row.write(5, 'N/A')

	try:
		with open('c:\\Challenge Compliance\\files\\user_manager.csv', 'rt') as f:
			manager = csv.reader(f, delimiter=',')
			for rows in manager:
				if data['db_list'][i]['owner']['uid'] == rows[1]:
					row.write(6, rows[3])
	except Exception as e:
		row.write(6, 'N/A')

	i+=1

try:
	wb.save('c:\\Challenge Compliance\\files\\db.xls')
except Exception as e:
	print(e)
