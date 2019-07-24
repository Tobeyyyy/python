from header import *
import xlwings as xw

def splitIntoSheet(input_file):
	print('opening ... ')
	wb = xw.Book(input_file)
	sht1 = wb.sheets[0]
	sheet_array={}
	for i in range(2,200):
		print('reading line '+str(i)+' ...')
		sheet_name=sht1.range('C'+str(i)).value
		date=sht1.range('B'+str(i)).value
		doc_number=sht1.range('D'+str(i)).value
		name=sht1.range('E'+str(i)).value
		amount=sht1.range('F'+str(i)).value
		status=sht1.range('G'+str(i)).value
		memo=sht1.range('H'+str(i)).value
		quantity=sht1.range('I'+str(i)).value
		location=sht1.range('J'+str(i)).value
		bin_number=sht1.range('K'+str(i)).value
		#seperate sheet by transaction type
		if sheet_name not in sheet_array.keys():
			sheet_array[sheet_name]=[[date,sheet_name,doc_number,name,amount,status,memo,quantity,location,bin_number]]
		else:
			sheet_array[sheet_name].append([date,sheet_name,doc_number,name,amount,status,memo,quantity,location,bin_number])
	print(sheet_array.keys())
	print('Spliting ... ')

	for key,value in sheet_array.items():
		if key != None:
			print('Spliting '+key+' ... ')
			total=0
			j=3
			invoice_sheet=wb.sheets.add(key)  
			current_sheet=wb.sheets[key]
			current_sheet.range('A1').value=['Date','Type','DocumentNumber','Name','Amount','Status','Memo','Qty','Location','BinNumber']
			for i in range(2,len(value)+2):
				current_sheet.range('A'+str(i)).value=value[i-2]
				j=j+1
				if value[i-2][7]=='':
					total=total+0
				else:
					total=total+value[i-2][7]
			current_sheet.range('H'+str(j)).value=total
			print('sheet '+key+' finishes')
			
	print('done ... ')
			
