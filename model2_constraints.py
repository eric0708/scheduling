import xlrd
import math

location = ("NumberOfShifts.xlsx")

workbook = xlrd.open_workbook(location)
sheet = workbook.sheet_by_index(0) 

location2 = ("peoplelabel_original.xlsx")

workbook2 = xlrd.open_workbook(location2)
sheet2 = workbook2.sheet_by_index(0)

f = open("number_of_shifts_output.txt", "r")
outputs = f.read()
outputs = outputs.replace('y', ' ')
outputs = outputs.replace('_', ' ')
outputlist = outputs.split()

numberofshiftsneeded = float(outputlist[0])

f.close()

f = open("model2_constraints.lp", "w")

weekends = (1, 7, 8, 14, 15, 21, 22, 28)

#count the total number of people
#organize people labels into different lists
	
numberofpeople = sheet2.nrows-1

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

for i in range(1, sheet2.nrows):
	if (sheet2.cell_value(i,1) == "高雄" and sheet2.cell_value(i,2) == "F"):
		KF = KF + [int(sheet2.cell_value(i,0))]
	elif(sheet2.cell_value(i,1) == "台北" and sheet2.cell_value(i,2) == "F"):
		TF = TF + [int(sheet2.cell_value(i,0))]
	elif(sheet2.cell_value(i,1) == "台中" and sheet2.cell_value(i,2) == "F"):
		CF = CF + [int(sheet2.cell_value(i,0))]
	elif(sheet2.cell_value(i,1) == "高雄" and sheet2.cell_value(i,2) == "M"):
		KM = KM + [int(sheet2.cell_value(i,0))]
	elif(sheet2.cell_value(i,1) == "台北" and sheet2.cell_value(i,2) == "M"):
		TM = TM + [int(sheet2.cell_value(i,0))]
	elif(sheet2.cell_value(i,1) == "台中" and sheet2.cell_value(i,2) == "M"):
		CM = CM + [int(sheet2.cell_value(i,0))]

F = F+KF+TF+CF
M = M+KM+TM+CM
K = K+KF+KM
T = T+TF+TM
C = C+CF+CM
All = All+F+M

lowerbound = input("Set lower bound(number between 0 and 1): ")
upperbound = input("Set upper bound(a non-negative integer): ")
consecutivedays = input("Limit of how many consecutive days(a positive integer): ")

#write objective function

message = ""

for i in All:
	for m in range(4):
		for k in range(2+7*m, 7+7*m):
			for j in range(1, 16):
				message = message + " + x"+str(i)+"_"+str(j)+"_"+str(k)
	for p in weekends:
		for q in range(16, 28):
			message = message + " + x"+str(i)+"_"+str(q)+"_"+str(p)
			
message = message+";\n"
message = message[2:]
message = "max:"+message

f.write(message)

#write max shifts constraints and give a lower bound for day shift people

for k in weekends:
		for j in range(16, 28):
			message = ""
			for i in All:
				message = message+" + x"+str(i)+"_"+str(j)+"_"+str(k)
			message = message+" <= "+str(math.ceil(1*float(sheet.cell_value(k,j)))+int(upperbound))+";\n"
			message = message[3:]
			f.write(message)		

for m in range(4):
	for k in range(2+7*m, 7+7*m):
		for j in range(1, 16):
			message = ""
			for i in All:
				message = message+" + x"+str(i)+"_"+str(j)+"_"+str(k)
			message = message+" <= "+str(math.ceil(1*float(sheet.cell_value(k,j)))+int(upperbound))+";\n"
			message = message[3:]
			f.write(message)

if (numberofshiftsneeded >= numberofpeople*20*1.3):
	for k in weekends:
			for j in range(16, 27):
				message = ""
				for i in All:
					message = message+" + x"+str(i)+"_"+str(j)+"_"+str(k)
				message = message+" >= "+str(math.ceil(float(lowerbound)*float(sheet.cell_value(k,j))))+";\n"
				message = message[3:]
				f.write(message)		

	for m in range(4):
		for k in range(2+7*m, 7+7*m):
			for j in range(1, 15):
				message = ""
				for i in All:
					message = message+" + x"+str(i)+"_"+str(j)+"_"+str(k)
				message = message+" >= "+str(math.ceil(float(lowerbound)*float(sheet.cell_value(k,j))))+";\n"
				message = message[3:]
				f.write(message)
else:
	for k in weekends:
			for j in range(16, 28):
				message = ""
				for i in All:
					message = message+" + x"+str(i)+"_"+str(j)+"_"+str(k)
				message = message+" >= "+str(math.ceil(float(lowerbound)*float(sheet.cell_value(k,j))))+";\n"
				message = message[3:]
				f.write(message)		

	for m in range(4):
		for k in range(2+7*m, 7+7*m):
			for j in range(1, 16):
				message = ""
				for i in All:
					message = message+" + x"+str(i)+"_"+str(j)+"_"+str(k)
				message = message+" >= "+str(math.ceil(float(lowerbound)*float(sheet.cell_value(k,j))))+";\n"
				message = message[3:]
				f.write(message)		
	

#write one shift per person per day constraints	
			
for i in All:
	for k in weekends:
		message = ""
		for j in range(16, 28):
			message = message+" + x"+str(i)+"_"+str(j)+"_"+str(k)
		message = message+" <= 1;\n"
		message = message[3:]
		f.write(message)
		
for i in All:
	for m in range(4):
		for k in range(2+7*m, 7+7*m):
			message = ""
			for j in range(1, 16):
				message = message+" + x"+str(i)+"_"+str(j)+"_"+str(k)
			message = message+" <= 1;\n"
			message = message[3:]
			f.write(message)		

#write at least one day off per week constraints
			
for i in All:
	for m in range(math.floor(28/float(consecutivedays))):
		message = ""
		for k in range(1+int(consecutivedays)*m, 8+int(consecutivedays)*m):
			if (k%7 == 1) or (k%7 == 0):
				for j in range(16, 28):
					message = message+" + x"+str(i)+"_"+str(j)+"_"+str(k)
			else:
				for n in range(1, 16):
					message = message+" + x"+str(i)+"_"+str(n)+"_"+str(k)
		message = message+" <= " + str(int(consecutivedays)-1) + ";\n"
		message = message[3:]
		f.write(message)

#write at least eight days off per month constraints
		
for i in All:
	message = ""
	for k in range(1, 29):
		if (k%7 == 1) or (k%7 == 0):
			for j in range(16, 28):
				message = message+" + x"+str(i)+"_"+str(j)+"_"+str(k)
		else:
			for n in range(1, 16):
				message = message+" + x"+str(i)+"_"+str(n)+"_"+str(k)
	message = message+" <= 20;\n"
	message = message[3:]
	f.write(message)
	
#write girls no night shifts constraint
	
message = ""

for i in F:
	for k in weekends:
		for j in range(24, 28):
			message = message + " + x"+str(i)+"_"+str(j)+"_"+str(k)
	for m in range(4):
		for n in range(2+7*m, 7+7*m):
			for p in range(12, 16):
				message	= message + " + x"+str(i)+"_"+str(p)+"_"+str(n)
message = message+" = 0;\n"
message = message[3:]
f.write(message)

#write taichung workers no weekend shifts constraint

message = ""

for k in weekends:
	for i in C:
		for j in range(16, 28):
			message = message+" + x"+str(i)+"_"+str(j)+"_"+str(k)
message = message+" = 0;\n"
message = message[3:]
f.write(message)

#Restrict the amount of people who works the night shift to within a range

if (numberofshiftsneeded >= numberofpeople*20*1.3):
	for j in range(24, 27):
		for k in weekends:
			if (int(sheet.cell_value(k,j)) > 0):
				message = ""
				for i in M:
					message = message + " + x"+str(i)+"_"+str(j)+"_"+str(k)
				message = message+" >= "+str(math.ceil(float(lowerbound)*float(sheet.cell_value(k,j))))+";\n"
				message = message[3:]
				f.write(message)	
			
	for j in range(12, 15):
		for m in range(4):
			for k in range(2+7*m, 7+7*m):
				if (int(sheet.cell_value(k,j)) > 0):
					message = ""
					for i in M:
						message = message + " + x"+str(i)+"_"+str(j)+"_"+str(k)
					message = message+" >= "+str(math.ceil(float(lowerbound)*float(sheet.cell_value(k,j))))+";\n"
					message = message[3:]
					f.write(message)
else:
	for j in range(24, 28):
		for k in weekends:
			if (int(sheet.cell_value(k,j)) > 0):
				message = ""
				for i in M:
					message = message + " + x"+str(i)+"_"+str(j)+"_"+str(k)
				message = message+" >= "+str(math.ceil(float(lowerbound)*float(sheet.cell_value(k,j))))+";\n"
				message = message[3:]
				f.write(message)	
			
	for j in range(12, 16):
		for m in range(4):
			for k in range(2+7*m, 7+7*m):
				if (int(sheet.cell_value(k,j)) > 0):
					message = ""
					for i in M:
						message = message + " + x"+str(i)+"_"+str(j)+"_"+str(k)
					message = message+" >= "+str(math.ceil(float(lowerbound)*float(sheet.cell_value(k,j))))+";\n"
					message = message[3:]
					f.write(message)
				
for j in range(24, 28):
	for k in weekends:
		if (int(sheet.cell_value(k,j)) > 0):
			message = ""
			for i in M:
				message = message + " + x"+str(i)+"_"+str(j)+"_"+str(k)
			message = message+" <= "+str(math.ceil(1*float(sheet.cell_value(k,j))))+";\n"
			message = message[3:]
			f.write(message)	
			
for j in range(12, 16):
	for m in range(4):
		for k in range(2+7*m, 7+7*m):
			if (int(sheet.cell_value(k,j)) > 0):
				message = ""
				for i in M:
					message = message + " + x"+str(i)+"_"+str(j)+"_"+str(k)
				message = message+" <= "+str(math.ceil(1*float(sheet.cell_value(k,j))))+";\n"
				message = message[3:]
				f.write(message)

#write take the day off before late night shifts constraints

for i in M:
	for k in weekends:
		if (k != 1) and (k%7 == 1):
			message = "x"+str(i)+"_27_"+str(k)
			for j in range(16, 28):
				message = message + " + x"+str(i)+"_"+str(j)+"_"+str(k-1)
			message = message+" <= 1;\n"
			f.write(message)
		elif (k%7 == 0):
			message = "x"+str(i)+"_27_"+str(k)
			for j in range(1, 16):
				message = message + " + x"+str(i)+"_"+str(j)+"_"+str(k-1)
			message = message+" <= 1;\n"
			f.write(message)
	for m in range(4):
		for k in range(2+7*m, 7+7*m):
			if (k%7 == 2):
				message = "x"+str(i)+"_15_"+str(k)
				for j in range(16, 28):
					message = message + " + x"+str(i)+"_"+str(j)+"_"+str(k-1)
				message = message+" <= 1;\n"
				f.write(message)
			else:
				message = "x"+str(i)+"_15_"+str(k)
				for j in range(1, 16):
					message = message + " + x"+str(i)+"_"+str(j)+"_"+str(k-1)
				message = message+" <= 1;\n"
				f.write(message)					

#write the variables used				

message = ""

for i in All:
	for m in range(4):
		for k in range(2+7*m, 7+7*m):
			for j in range(1, 16):
				message = message + ", x"+str(i)+"_"+str(j)+"_"+str(k)
	for p in weekends:
		for q in range(16, 28):
			message = message + ", x"+str(i)+"_"+str(q)+"_"+str(p)
			
message = message+";\n"
message = message[1:]
message = "bin"+message

f.write(message)