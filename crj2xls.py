#!python3
import csv
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
import os
from datetime import datetime
import glob

start_file = max(glob.glob('*.csv'),key=os.path.getctime)

with open(start_file,'r') as infile:
	csv = csv.reader(infile,delimiter=',')
	records = {}
	for rownum,row in enumerate(csv):
		if rownum == 0:
			headers = [item.strip() for item in row]
			continue
		if row[headers.index('CLIENT')].strip()+' '+row[headers.index('CL-NUM')].strip() not in records:
			records[row[headers.index('CLIENT')].strip()+' '+row[headers.index('CL-NUM')].strip()] = {}
		if row[headers.index('CRJ#')] not in  records[row[headers.index('CLIENT')].strip()+' '+row[headers.index('CL-NUM')].strip()]:
			records[row[headers.index('CLIENT')].strip()+' '+row[headers.index('CL-NUM')].strip()][row[headers.index('CRJ#')]] = {'COURT COSTS':0,'FIRM MONEY':0}
		try:
			records[row[headers.index('CLIENT')].strip()+' '+row[headers.index('CL-NUM')].strip()][row[headers.index('CRJ#')]]['COURT COSTS'] += float(row[headers.index('COURT COST')])
		except:
			print((row[headers.index('COURT COST')],))
			break
		try:
			records[row[headers.index('CLIENT')].strip()+' '+row[headers.index('CL-NUM')].strip()][row[headers.index('CRJ#')]]['FIRM MONEY'] += float(row[headers.index('FIRM MONEY')])
		except:
			print((row[headers.index('FIRM MONEY')],))
			break

datetimestamp = datetime.now().strftime('%Y%m%d%H%M%S')
filename = 'A10_Compile_'+datetimestamp+'.xlsx'
with xlsxwriter.Workbook(os.path.join('output',filename)) as wb:
	ws = wb.add_worksheet()
	grey_header = wb.add_format({'bg_color':'#808080','font_color':'#FFFFFF','bold':True,'align':'center'})
	currency = wb.add_format({'num_format':'_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'})
	bold_num = wb.add_format({'bold':True,'num_format':'_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'})
	bold = wb.add_format({'bold':True})
	ws.merge_range(0,2,0,3,'COST',grey_header)
	ws.merge_range(0,4,0,5,'FEES',grey_header)
	head_row = ["STATUS","CLIENT / CLIENT #","BATCH #"," AMT PAID ","BATCH #"," AMT PAID ","CLIENT'S CHECK #"]
	ws.write_row(1,0,head_row,grey_header)

	rownum = 2
	for m9,batch_dict in records.items():
		for batch_num,amt_dict in batch_dict.items():
			ws.write(rownum,1,m9)
			if amt_dict['COURT COSTS'] > 0:
				ws.write(rownum,2,batch_num)
				ws.write(rownum,3,amt_dict['COURT COSTS'],currency)
			if amt_dict['FIRM MONEY'] > 0:
				ws.write(rownum,4,batch_num)
				ws.write(rownum,5,amt_dict['FIRM MONEY'],currency)
			rownum+=1

	sum_costs = '=SUM({}:{})'.format(xl_rowcol_to_cell(2,3),xl_rowcol_to_cell(rownum-1,3))
	sum_fees = '=SUM({}:{})'.format(xl_rowcol_to_cell(2,5),xl_rowcol_to_cell(rownum-1,5))

	ws.write(rownum,0,'Totals:',bold)
	ws.write_formula(rownum,3,sum_costs,bold_num)
	ws.write_formula(rownum,5,sum_fees,bold_num)

	column_widths = [11.71,59,17,20.43,17,17,56.71]
	for colnum,width in enumerate(column_widths):
		ws.set_column(first_col = colnum,last_col = colnum,width = width)

print("Output file completed and in output folder: "+filename)
os.system('pause')
