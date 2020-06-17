import xlrd
import numpy as np
import xlsxwriter
from datetime import date

#transfer into results for evaluation

location1 = ("schedule.xlsx")

workbook1 = xlrd.open_workbook(location1)
sheet1 = workbook1.sheet_by_index(0) 


location = ("FinalShifts.xlsx")

workbook = xlrd.open_workbook(location)
sheet = workbook.sheet_by_index(0) 

numberofpeople = sheet.ncols-1

outputlist = [0 for x in range(numberofpeople*28)]

for i in range(1, 29):
	for j in range(1, numberofpeople+1):
		outputlist[numberofpeople*(i-1)+(j-1)] = sheet.cell_value(i,j)

chart = np.zeros((29, 48))

for i in range(numberofpeople*28):
	if outputlist[i] != '':
		if int(outputlist[i]) == 1:
			for m in range(13, 30):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			for n in range(21, 22):
				chart[int(i/numberofpeople),n] = chart[int(i/numberofpeople),n]-1
		elif int(outputlist[i]) == 2:
			for m in range(15, 32):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			for n in range(23, 24):
				chart[int(i/numberofpeople),n] = chart[int(i/numberofpeople),n]-1
		elif int(outputlist[i]) == 3:
			for m in range(16, 34):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			for n in range(22, 24):
				chart[int(i/numberofpeople),n] = chart[int(i/numberofpeople),n]-1
		elif int(outputlist[i]) == 4:
			for m in range(16, 34):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			for n in range(24, 26):
				chart[int(i/numberofpeople),n] = chart[int(i/numberofpeople),n]-1		
		elif int(outputlist[i]) == 5:
			for m in range(17, 35):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			for n in range(24, 26):
				chart[int(i/numberofpeople),n] = chart[int(i/numberofpeople),n]-1
		elif int(outputlist[i]) == 6:
			for m in range(18, 36):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			for n in range(25, 27):
				chart[int(i/numberofpeople),n] = chart[int(i/numberofpeople),n]-1	
		elif int(outputlist[i]) == 7:
			for m in range(19, 37):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			for n in range(26, 28):
				chart[int(i/numberofpeople),n] = chart[int(i/numberofpeople),n]-1
		elif int(outputlist[i]) == 8:
			for m in range(20, 38):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			for n in range(28, 30):
				chart[int(i/numberofpeople),n] = chart[int(i/numberofpeople),n]-1	
		elif int(outputlist[i]) == 9:
			for m in range(24, 41):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			for n in range(34, 35):
				chart[int(i/numberofpeople),n] = chart[int(i/numberofpeople),n]-1
		elif int(outputlist[i]) == 10:
			for m in range(27, 44):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			for n in range(34, 35):
				chart[int(i/numberofpeople),n] = chart[int(i/numberofpeople),n]-1		
		elif int(outputlist[i]) == 11:
			for m in range(27, 44):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			for n in range(35, 36):
				chart[int(i/numberofpeople),n] = chart[int(i/numberofpeople),n]-1
		elif int(outputlist[i]) == 12:
			for m in range(29, 46):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			for n in range(35, 36):
				chart[int(i/numberofpeople),n] = chart[int(i/numberofpeople),n]-1	
		elif int(outputlist[i]) == 13:
			for m in range(31, 48):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			chart[int(i/numberofpeople),39] = chart[int(i/numberofpeople),39]-1
		elif int(outputlist[i]) == 14:
			for m in range(33, 48):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			chart[int(i/numberofpeople),41] = chart[int(i/numberofpeople),41]-1
			chart[int(i/numberofpeople)+1,0] = chart[int(i/numberofpeople)+1,0]+1
			chart[int(i/numberofpeople)+1,1] = chart[int(i/numberofpeople)+1,1]+1
		elif int(outputlist[i]) == 15:
			for m in range(1, 17):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
		elif int(outputlist[i]) == 16:
			for m in range(15, 32):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			chart[int(i/numberofpeople),23] = chart[int(i/numberofpeople),23]-1
		elif int(outputlist[i]) == 17:
			for m in range(16, 33):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			chart[int(i/numberofpeople),24] = chart[int(i/numberofpeople),24]-1
		elif int(outputlist[i]) == 18:
			for m in range(16, 33):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			chart[int(i/numberofpeople),25] = chart[int(i/numberofpeople),25]-1
		elif int(outputlist[i]) == 19:
			for m in range(18, 35):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			chart[int(i/numberofpeople),24] = chart[int(i/numberofpeople),24]-1
		elif int(outputlist[i]) == 20:
			for m in range(18, 35):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			chart[int(i/numberofpeople),25] = chart[int(i/numberofpeople),25]-1	
		elif int(outputlist[i]) == 21:
			for m in range(20, 37):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			chart[int(i/numberofpeople),28] = chart[int(i/numberofpeople),28]-1
		elif int(outputlist[i]) == 22:
			for m in range(27, 44):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			chart[int(i/numberofpeople),34] = chart[int(i/numberofpeople),34]-1
		elif int(outputlist[i]) == 23:
			for m in range(27, 44):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			chart[int(i/numberofpeople),35] = chart[int(i/numberofpeople),35]-1
		elif int(outputlist[i]) == 24:
			for m in range(29, 46):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			chart[int(i/numberofpeople),35] = chart[int(i/numberofpeople),35]-1
		elif int(outputlist[i]) == 25:
			for m in range(31, 48):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			chart[int(i/numberofpeople),39] = chart[int(i/numberofpeople),39]-1
		elif int(outputlist[i]) == 26:
			for m in range(33, 48):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1
			chart[int(i/numberofpeople),41] = chart[int(i/numberofpeople),41]-1
			chart[int(i/numberofpeople)+1,0] = chart[int(i/numberofpeople)+1,0]+1
			chart[int(i/numberofpeople)+1,1] = chart[int(i/numberofpeople)+1,1]+1
		elif int(outputlist[i]) == 27:
			for m in range(1, 17):
				chart[int(i/numberofpeople),m] = chart[int(i/numberofpeople),m]+1		

workbook = xlsxwriter.Workbook('Results.xlsx')
worksheet = workbook.add_worksheet()	

for i in range(1, 16):
	date_value = xlrd.xldate_as_tuple(sheet1.cell_value(i,0),workbook1.datemode)
	worksheet.write(i, 0, date(*date_value[:3]).strftime('%Y/%m/%d'))
	
for i in range(16, 29):
	date_value = xlrd.xldate_as_tuple(sheet1.cell_value(i,0),workbook1.datemode)
	worksheet.write(i, 0, date(*date_value[:3]).strftime('%Y/%m/%d'))

for i in range(1, 49):
	if i%2 == 1:
		timeoutput = str(int(i/2))+":00"
	else:
		timeoutput = str(int(i/2)-1)+":30"
	worksheet.write(0, i, timeoutput)
				
for i in range(1, 29):
	for j in range(1, 49):
		worksheet.write(i, j, chart[i-1, j-1])

workbook.close()