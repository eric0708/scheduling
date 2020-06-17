import xlsxwriter
import xlrd
from datetime import date

#sort the shifts from small to large

location1 = ("Shifts.xlsx")

workbook1 = xlrd.open_workbook(location1)
worksheet1 = workbook1.sheet_by_index(0) 

workbook2 = xlsxwriter.Workbook('FinalShifts.xlsx')
worksheet2 = workbook2.add_worksheet()

location3 = ("peoplelabel_original.xlsx")

workbook3 = xlrd.open_workbook(location3)
sheet3 = workbook3.sheet_by_index(0)

location4 = ("schedule.xlsx")

workbook4 = xlrd.open_workbook(location4)
sheet4 = workbook4.sheet_by_index(0) 

numberofpeople = sheet3.nrows-1

KF = []
TF = []
CF = []
KM = []
TM = []
CM = []
F = []
M = []
K = []
T = []
C = []
All = []

for i in range(1, sheet3.nrows):
	if (sheet3.cell_value(i,1) == "高雄" and sheet3.cell_value(i,2) == "F"):
		KF = KF + [int(sheet3.cell_value(i,0))]
	elif(sheet3.cell_value(i,1) == "台北" and sheet3.cell_value(i,2) == "F"):
		TF = TF + [int(sheet3.cell_value(i,0))]
	elif(sheet3.cell_value(i,1) == "台中" and sheet3.cell_value(i,2) == "F"):
		CF = CF + [int(sheet3.cell_value(i,0))]
	elif(sheet3.cell_value(i,1) == "高雄" and sheet3.cell_value(i,2) == "M"):
		KM = KM + [int(sheet3.cell_value(i,0))]
	elif(sheet3.cell_value(i,1) == "台北" and sheet3.cell_value(i,2) == "M"):
		TM = TM + [int(sheet3.cell_value(i,0))]
	elif(sheet3.cell_value(i,1) == "台中" and sheet3.cell_value(i,2) == "M"):
		CM = CM + [int(sheet3.cell_value(i,0))]

F = F+KF+TF+CF
M = M+KM+TM+CM
K = K+KF+KM
T = T+TF+TM
C = C+CF+CM
All = All+F+M


for k in range(1, 29):
	girlindex = []
	girlshifts = []
	boyindex = []
	boyshifts = []
	boynightindex = []
	boynightshifts = []
	
	for i in F:
		if (worksheet1.cell_value(k, i) != ''):
			girlindex = girlindex + [i]
			girlshifts = girlshifts + [worksheet1.cell_value(k, i)]
	girlshifts.sort()
	for m in range(len(girlindex)):
		worksheet2.write(k, girlindex[m], girlshifts[m])
	
	
	
	for i in M:
		if (worksheet1.cell_value(k, i) != '') and (worksheet1.cell_value(k, i) != 15) and (worksheet1.cell_value(k, i) != 27):
			boyindex = boyindex + [i]
			boyshifts = boyshifts + [worksheet1.cell_value(k, i)]
		elif (worksheet1.cell_value(k, i) != 15) or (worksheet1.cell_value(k, i) != 27):
			boynightindex = boynightindex + [i]
			boynightshifts = boynightshifts + [worksheet1.cell_value(k, i)]
	boyshifts.sort()
	for m in range(len(boyindex)):
		worksheet2.write(k, boyindex[m], boyshifts[m])
	for m in range(len(boynightindex)):
		worksheet2.write(k, boynightindex[m], boynightshifts[m])
		
for i in All:
	worksheet2.write(0, i, worksheet1.cell_value(0, i))

for i in range(1, 16):
	date_value = xlrd.xldate_as_tuple(sheet4.cell_value(i,0),workbook4.datemode)
	worksheet2.write(i, 0, date(*date_value[:3]).strftime('%Y/%m/%d'))
	
for i in range(16, 29):
	date_value = xlrd.xldate_as_tuple(sheet4.cell_value(i,0),workbook4.datemode)
	worksheet2.write(i, 0, date(*date_value[:3]).strftime('%Y/%m/%d'))

workbook2.close()