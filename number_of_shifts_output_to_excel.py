import xlrd
import xlsxwriter
import math
from datetime import date

#writing the output into an excel file

f = open("number_of_shifts_output.txt", "r")
outputs = f.read()
outputs = outputs.replace('y', ' ')
outputs = outputs.replace('_', ' ')
outputlist = outputs.split()

workbook = xlsxwriter.Workbook('NumberOfShifts.xlsx')
worksheet = workbook.add_worksheet()

location2 = ("schedule.xlsx")

workbook2 = xlrd.open_workbook(location2)
sheet2 = workbook2.sheet_by_index(0) 


location1 = ("peoplelabel_organized.xlsx")

workbook1 = xlrd.open_workbook(location1)
worksheet1 = workbook1.sheet_by_index(0) 

numberofpeople = worksheet1.nrows-1

numberofshiftsneeded = float(outputlist[0])

lowerbound = (numberofpeople*20*0.7)/numberofshiftsneeded
preferredlowerbound = math.ceil(lowerbound*100)/100
print("Preferred lower bound: ", preferredlowerbound)

if(numberofshiftsneeded >= numberofpeople*20*1.3):
	preferredupperbound = 0
else:
	preferredupperbound = math.ceil((numberofpeople*20*1.3-numberofshiftsneeded)/(8*8+11*20))
print("Preferred upper bound: ", preferredupperbound)

for i in range(1, 28):
	worksheet.write(0, i, i)

for i in range(1, 16):
	date_value = xlrd.xldate_as_tuple(sheet2.cell_value(i,0),workbook2.datemode)
	worksheet.write(i, 0, date(*date_value[:3]).strftime('%Y/%m/%d'))
	
for i in range(16, 29):
	date_value = xlrd.xldate_as_tuple(sheet2.cell_value(i,0),workbook2.datemode)
	worksheet.write(i, 0, date(*date_value[:3]).strftime('%Y/%m/%d'))

for i in range(0, 396):
	worksheet.write(int(outputlist[3*i+2]), int(outputlist[3*i+1]), int(outputlist[3*i+3]))
	
workbook.close()