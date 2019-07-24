from header import *

def checkNewFile(input_folder):#get the lastest file
	list_of_files = glob.glob(input_folder+'\\*.csv')
	latest_file = max(list_of_files, key=os.path.getctime)
	print(latest_file)
	return latest_file
#input_file=checkNewFile(input_folder)
input_file='C:\\Users\\melody\\AppData\\Local\\Programs\\Python\\Python36-32\\mytools\\transactionCheck\\CheckItemAllTransactionResults55.csv'

def readFile(input_file):
	all_type={}
	with open(input_file,newline='') as inputfile:
		readfile=csv.reader(inputfile);
		for row in readfile:
			if row[0] != 'Internal ID':
				#name_list.append([row[0]]);
				type_name=row[2]
				qty=row[8]
				if type_name=='Item Receipt' and row[4]=='' and int(qty)>0:
					if 'Transfer Receipt' not in all_type.keys():
						all_type['Transfer Receipt']=qty
					else:
						all_type['Transfer Receipt']=int(all_type['Transfer Receipt'])+int(qty)
				if type_name=='Item Fulfillment' and row[4]=='':
					if 'Transfer Fulfillment' not in all_type.keys():
						all_type['Transfer Fulfillment']=qty
					else:
						all_type['Transfer Fulfillment']=int(all_type['Transfer Fulfillment'])+int(qty)
				if type_name=='Inventory Adjustment' and row[10]=='LI':
					if 'LI Inventory Adjustment' not in all_type.keys():
						all_type['LI Inventory Adjustment']=qty
					else:
						all_type['LI Inventory Adjustment']=int(all_type['LI Inventory Adjustment'])+int(qty)
				if type_name not in all_type.keys():
					all_type[type_name]=qty
				else:
					all_type[type_name]=int(all_type[type_name])+int(qty)

			else:
				header=row;#save header of the file
#	print("Success: open file",all_type)
					
	if 'LI Inventory Adjustment' not in all_type.keys():
		all_type['LI Inventory Adjustment']=0
	return(all_type)

def wirteFile(result):
	result_array=[]
	output_name='C:\\Users\\melody\\AppData\\Local\\Programs\\Python\\Python36-32\\mytools\\transactionCheck\\transaction_result.xlsx'
	wb = xw.Book(output_name)
	sht1 = wb.sheets['Sheet1']
	i=1
	sht1.range('A1').value='Type'
	sht1.range('B1').value='Qty'
	for key,value in result.items():
		i=i+1
		print(str(key)+':'+str(value))
		result_array.append([key,value])
		sht1.range('A'+str(i)).value=key
		sht1.range('B'+str(i)).value=value
	salesorder_fulfillment=sht1.range('B4').value-sht1.range('B7').value
	if int(salesorder_fulfillment)==int(sht1.range('B5').value):
		print('fulfillment qty is correct'+str(salesorder_fulfillment))
	else:
		print('actual sales order fulfillment= '+str(salesorder_fulfillment))
	
	wb.save()	
		
#result=readFile(input_file)
#wirteFile(result)
fileName='CheckItemAllTransactionResults.xls'
splitIntoSheet(input_folder+fileName)