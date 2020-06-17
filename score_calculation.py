import xlrd

#calculate evaluation results

location1 = ("Results.xlsx")

workbook1 = xlrd.open_workbook(location1)
sheet1 = workbook1.sheet_by_index(0) 

location2 = ("schedule.xlsx")

workbook2 = xlrd.open_workbook(location2)
sheet2 = workbook2.sheet_by_index(0) 

location3 = ("callspredict.xlsx")

workbook3 = xlrd.open_workbook(location3)
sheet3 = workbook3.sheet_by_index(0) 

location4 = ("peoplelabel_organized.xlsx")

workbook4 = xlrd.open_workbook(location4)
sheet4 = workbook4.sheet_by_index(0) 

location5 = ("FinalShifts.xlsx")

workbook5 = xlrd.open_workbook(location5)
sheet5 = workbook5.sheet_by_index(0) 

f = open("EvaluationResults.txt", "w")

giveuprate = []

for i in range(1, 29):
	if (sheet1.cell_value(i,1) <= sheet2.cell_value(i,1)):
		giveuprate = giveuprate+[(1-((float(sheet2.cell_value(i,1))-float(sheet1.cell_value(i,1)))/float(sheet2.cell_value(i,1))))*100*0.7]
	else:
		giveuprate = giveuprate+[100*0.7]
	
	for j in range(18, 38):
		if (sheet1.cell_value(i,j) <= sheet2.cell_value(i,j)):
			giveuprate = giveuprate+[(1-((float(sheet2.cell_value(i,j))-float(sheet1.cell_value(i,j)))/float(sheet2.cell_value(i,j))))*100]
		else:
			giveuprate = giveuprate+[100]
	for j in range(38, 49):
		if (sheet1.cell_value(i,j) <= sheet2.cell_value(i,j)):
			giveuprate = giveuprate+[(1-((float(sheet2.cell_value(i,j))-float(sheet1.cell_value(i,j)))/float(sheet2.cell_value(i,j))))*100*0.7]
		else:
			giveuprate = giveuprate+[100*0.7]
	for j in range(2, 18):
		if (sheet1.cell_value(i,j) <= sheet2.cell_value(i,j)):
			giveuprate = giveuprate+[(1-((float(sheet2.cell_value(i,j))-float(sheet1.cell_value(i,j)))/float(sheet2.cell_value(i,j))))*100*0.2]
		else:
			giveuprate = giveuprate+[100*0.2]

sumgiveuprate = sum(giveuprate)
giveupratescore = (sumgiveuprate/88480)
message = "Give Up Rate Score: " + str(giveupratescore) + "\n"
print(message)
f.write(message)

servicerate = []

for i in range(1, 29):
	service = 70.5838-0.4091*float(sheet3.cell_value(i,1))+1.3143*float(sheet1.cell_value(i,1))
	if (service < 0):
		service = 0
	if (service > 100):
		service = 100
	service = service*0.7
	servicerate = servicerate+[service]
	
	for j in range(18, 38):
		service = 70.5838-0.4091*float(sheet3.cell_value(i,j))+1.3143*float(sheet1.cell_value(i,j))
		if (service < 0):
			service = 0
		if (service > 100):
			service = 100
		servicerate = servicerate+[service]
	for j in range(38, 49):
		service = 70.5838-0.4091*float(sheet3.cell_value(i,j))+1.3143*float(sheet1.cell_value(i,j))
		if (service < 0):
			service = 0
		if (service > 100):
			service = 100
		service = service*0.7
		servicerate = servicerate+[service]
	for j in range(2, 18):
		service = 70.5838-0.4091*float(sheet3.cell_value(i,j))+1.3143*float(sheet1.cell_value(i,j))
		if (service < 0):
			service = 0
		if (service > 100):
			service = 100
		service = service*0.2
		servicerate = servicerate+[service]

sumservicerate = sum(servicerate)
serviceratescore = (sumservicerate/88480)
message = "Service Rate Score: " + str(serviceratescore) + "\n"
print(message)
f.write(message)

actualcalls = []
predictcalls = []

for i in range(1, 29):
	
	if (sheet1.cell_value(i,1) <= sheet2.cell_value(i,1)):
		actualcalls = actualcalls+[float(sheet1.cell_value(i,1))*2.5*0.7]
	else:
		actualcalls = actualcalls+[float(sheet2.cell_value(i,1))*2.5*0.7]
	predictcalls = predictcalls + [float(sheet2.cell_value(i,1))*2.5*0.7]

	for j in range(18, 38):
		if (sheet1.cell_value(i,j) <= sheet2.cell_value(i,j)):
			actualcalls = actualcalls+[float(sheet1.cell_value(i,j))*2.5]
		else:
			actualcalls = actualcalls+[float(sheet2.cell_value(i,j))*2.5]
		predictcalls = predictcalls + [float(sheet2.cell_value(i,j))*2.5]
	for j in range(38, 49):
		if (sheet1.cell_value(i,j) <= sheet2.cell_value(i,j)):
			actualcalls = actualcalls+[float(sheet1.cell_value(i,j))*2.5*0.7]
		else:
			actualcalls = actualcalls+[float(sheet2.cell_value(i,j))*2.5*0.7]
		predictcalls = predictcalls + [float(sheet2.cell_value(i,j))*2.5*0.7]
	for j in range(2, 18):
		if (sheet1.cell_value(i,j) <= sheet2.cell_value(i,j)):
			actualcalls = actualcalls+[float(sheet1.cell_value(i,j))*2.5*0.2]
		else:
			actualcalls = actualcalls+[float(sheet2.cell_value(i,j))*2.5*0.2]
		predictcalls = predictcalls + [float(sheet2.cell_value(i,j))*2.5*0.2]

sumactual = sum(actualcalls)
sumpredict = sum(predictcalls)
callscore = (sumactual/sumpredict)	
message = "Call Score: " + str(callscore) + "\n"
print(message)
f.write(message)

counter = 0

for i in range(1, 29):
	for j in range(1, 49):
		if (float(sheet2.cell_value(i,j)) != 0):
			if (float(sheet1.cell_value(i,j))/float(sheet2.cell_value(i,j)) >= 0.7):
				counter = counter + 1
			"""
			else:
				message = str(i)+"   "+str(j)+"   "+str(float(sheet1.cell_value(i,j))/float(sheet2.cell_value(i,j)))+"\n" 
				print(message)
			"""
		
productivityscore = (float(counter)/1344.0)
message = "Productivity Score: " + str(productivityscore) + "\n"
print(message)
f.write(message)

counter = 0

for g in range(1,8):
	for m in range(4):
		for k in range(2+7*m, 7+7*m):
			shifts = []
			counts = [0]*13
			for i in range(1, sheet4.nrows):
				if (sheet4.cell_value(i,3) == "K"+str(g)):
					if (sheet5.cell_value(k,i) != ''):
						shifts = shifts+[int(sheet5.cell_value(k,i))]
			if (len(shifts) != 0):
				for i in range(len(shifts)):
					if(shifts[i] == 1):
						counts[0] += 1
					elif(shifts[i] == 2):
						counts[1] += 1
					elif((shifts[i] == 3) or (shifts[i] == 4)):
						counts[2] += 1
					elif(shifts[i] == 5):
						counts[3] += 1
					elif(shifts[i] == 6):
						counts[4] += 1
					elif(shifts[i] == 7):
						counts[5] += 1
					elif(shifts[i] == 8):
						counts[6] += 1
					elif(shifts[i] == 9):
						counts[7] += 1
					elif((shifts[i] == 10) or (shifts[i] == 11)):
						counts[8] += 1
					elif(shifts[i] == 12):
						counts[9] += 1
					elif(shifts[i] == 13):
						counts[10] += 1
					elif(shifts[i] == 14):
						counts[11] += 1
					elif(shifts[i] == 15):
						counts[12] += 1
				maxshift = max(counts)
				maxindex = counts.index(maxshift)
				if (maxindex == 0):
					if(counts[12]>counts[1]):
						finalmax = maxshift + counts[12]
					else:
						finalmax = maxshift + counts[1]
				elif (maxindex == 12):
						finalmax = maxshift + counts[11]
				else:
					if(counts[maxindex+1]>counts[maxindex-1]):
						finalmax = maxshift + counts[maxindex+1]
					else:
						finalmax = maxshift + counts[maxindex-1]
				if(float(finalmax)/len(shifts) >= 0.7):
					counter += 1

for g in range(1,4):
	for m in range(4):
		for k in range(2+7*m, 7+7*m):
			shifts = []
			counts = [0]*13
			for i in range(1, sheet4.nrows):
				if (sheet4.cell_value(i,3) == "T"+str(g)):
					if (sheet5.cell_value(k,i) != ''):
						shifts = shifts+[int(sheet5.cell_value(k,i))]
			if (len(shifts) != 0):
				for i in range(len(shifts)):
					if(shifts[i] == 1):
						counts[0] += 1
					elif(shifts[i] == 2):
						counts[1] += 1
					elif((shifts[i] == 3) or (shifts[i] == 4)):
						counts[2] += 1
					elif(shifts[i] == 5):
						counts[3] += 1
					elif(shifts[i] == 6):
						counts[4] += 1
					elif(shifts[i] == 7):
						counts[5] += 1
					elif(shifts[i] == 8):
						counts[6] += 1
					elif(shifts[i] == 9):
						counts[7] += 1
					elif((shifts[i] == 10) or (shifts[i] == 11)):
						counts[8] += 1
					elif(shifts[i] == 12):
						counts[9] += 1
					elif(shifts[i] == 13):
						counts[10] += 1
					elif(shifts[i] == 14):
						counts[11] += 1
					elif(shifts[i] == 15):
						counts[12] += 1
				maxshift = max(counts)
				maxindex = counts.index(maxshift)
				if (maxindex == 0):
					if(counts[12]>counts[1]):
						finalmax = maxshift + counts[12]
					else:
						finalmax = maxshift + counts[1]
				elif (maxindex == 12):
						finalmax = maxshift + counts[11]
				else:
					if(counts[maxindex+1]>counts[maxindex-1]):
						finalmax = maxshift + counts[maxindex+1]
					else:
						finalmax = maxshift + counts[maxindex-1]
				if(float(finalmax)/len(shifts) >= 0.7):
					counter += 1		
				
for g in range(1,3):
	for m in range(4):
		for k in range(2+7*m, 7+7*m):
			shifts = []
			counts = [0]*13
			for i in range(1, sheet4.nrows):
				if (sheet4.cell_value(i,3) == "C"+str(g)):
					if (sheet5.cell_value(k,i) != ''):
						shifts = shifts+[int(sheet5.cell_value(k,i))]
			if (len(shifts) != 0):
				for i in range(len(shifts)):
					if(shifts[i] == 1):
						counts[0] += 1
					elif(shifts[i] == 2):
						counts[1] += 1
					elif((shifts[i] == 3) or (shifts[i] == 4)):
						counts[2] += 1
					elif(shifts[i] == 5):
						counts[3] += 1
					elif(shifts[i] == 6):
						counts[4] += 1
					elif(shifts[i] == 7):
						counts[5] += 1
					elif(shifts[i] == 8):
						counts[6] += 1
					elif(shifts[i] == 9):
						counts[7] += 1
					elif((shifts[i] == 10) or (shifts[i] == 11)):
						counts[8] += 1
					elif(shifts[i] == 12):
						counts[9] += 1
					elif(shifts[i] == 13):
						counts[10] += 1
					elif(shifts[i] == 14):
						counts[11] += 1
					elif(shifts[i] == 15):
						counts[12] += 1
				maxshift = max(counts)
				maxindex = counts.index(maxshift)
				if (maxindex == 0):
					if(counts[12]>counts[1]):
						finalmax = maxshift + counts[12]
					else:
						finalmax = maxshift + counts[1]
				elif (maxindex == 12):
						finalmax = maxshift + counts[11]
				else:
					if(counts[maxindex+1]>counts[maxindex-1]):
						finalmax = maxshift + counts[maxindex+1]
					else:
						finalmax = maxshift + counts[maxindex-1]
				if(float(finalmax)/len(shifts) >= 0.7):
					counter += 1

groupmeetingscore = float(counter)/240
message = "Group Meeting Score: " + str(groupmeetingscore) + "\n"
print(message)
f.write(message)	