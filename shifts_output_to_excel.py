import xlsxwriter
import xlrd
from datetime import date

#writing the shifts output into an excel file

f = open("shifts_output.txt", "r")
outputs = f.read()
outputs = outputs.replace('x', ' ')
outputs = outputs.replace('_', ' ')
outputlist = outputs.split()

location2 = ("schedule.xlsx")

workbook2 = xlrd.open_workbook(location2)
sheet2 = workbook2.sheet_by_index(0) 

workbook = xlsxwriter.Workbook('Shifts.xlsx')
worksheet = workbook.add_worksheet()

location1 = ("peoplelabel_organized.xlsx")

workbook1 = xlrd.open_workbook(location1)
worksheet1 = workbook1.sheet_by_index(0) 

count = len(open("shifts_output.txt").readlines())

numberofpeople = worksheet1.nrows-1

for i in range(1, worksheet1.nrows):
	worksheet.write(0, i, worksheet1.cell_value(i, 0))

for i in range(1, 16):
	date_value = xlrd.xldate_as_tuple(sheet2.cell_value(i,0),workbook2.datemode)
	worksheet.write(i, 0, date(*date_value[:3]).strftime('%Y/%m/%d'))
	
for i in range(16, 29):
	date_value = xlrd.xldate_as_tuple(sheet2.cell_value(i,0),workbook2.datemode)
	worksheet.write(i, 0, date(*date_value[:3]).strftime('%Y/%m/%d'))

for i in range(0, count):
	if int(outputlist[4*i+3]) == 1:
		worksheet.write(int(outputlist[4*i+2]), int(outputlist[4*i]), int(outputlist[4*i+1]))

workbook.close()