import xlrd
import math

coverage = 1

location = ("schedule.xlsx")

workbook = xlrd.open_workbook(location)
sheet = workbook.sheet_by_index(0) 

f = open("model1_constraints.lp", "w")

#write objective function

f.write("min: ")

message = ""

for i in range(4):
	for j in range(2+7*i, 7+7*i):
		for k in range(1, 16):
			message = message+" + y"+str(k)+"_"+str(j)
			
weekends = (1, 7, 8, 14, 15, 21, 22, 28)

for i in weekends:
	for k in range(16, 28):
		message = message+" + y"+str(k)+"_"+str(i)
		
message = message+";\n"
message = message[3:]

#write weekday constraints

f.write(message)

for j in range(4):
	for i in range(2+7*j, 7+7*j):
		if i%7 == 2: 
			#1
			message = "y26_"+str(i-1)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,1)))+";\n"
			f.write(message)
			#2
			message = "y26_"+str(i-1)+" + y15_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,2)))+";\n"
			f.write(message)
		else:
			#1
			message = "y14_"+str(i-1)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,1)))+";\n"
			f.write(message)
			#2
			message = "y14_"+str(i-1)+" + y15_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,2)))+";\n"
			f.write(message)
		#3
		message = "y15_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,3)))+";\n"
		f.write(message)
		#4
		message = "y15_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,4)))+";\n"
		f.write(message)
		#5
		message = "y15_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,5)))+";\n"
		f.write(message)
		#6
		message = "y15_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,6)))+";\n"
		f.write(message)
		#7
		message = "y15_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,7)))+";\n"
		f.write(message)
		#8
		message = "y15_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,8)))+";\n"
		f.write(message)
		#9
		message = "y15_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,9)))+";\n"
		f.write(message)
		#10
		message = "y15_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,10)))+";\n"
		f.write(message)
		#11
		message = "y15_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,11)))+";\n"
		f.write(message)
		#12
		message = "y15_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,12)))+";\n"
		f.write(message)
		#13
		message = "y15_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,13)))+";\n"
		f.write(message)
		#14
		message = "y15_"+str(i)+" + y1_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,14)))+";\n"
		f.write(message)
		#15
		message = "y15_"+str(i)+" + y1_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,15)))+";\n"
		f.write(message)
		#16
		message = "y15_"+str(i)+" + y1_"+str(i)+" + y2_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,16)))+";\n"
		f.write(message)
		#17
		message = "y15_"+str(i)+" + y1_"+str(i)+" + y2_"+str(i)+" + y3_"+str(i)+" + y4_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,17)))+";\n"
		f.write(message)
		#18
		message = "y1_"+str(i)+" + y2_"+str(i)+" + y3_"+str(i)+" + y4_"+str(i)+" + y5_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,18)))+";\n"
		f.write(message)
		#19
		message = "y1_"+str(i)+" + y2_"+str(i)+" + y3_"+str(i)+" + y4_"+str(i)+" + y5_"+str(i)+" + y6_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,19)))+";\n"
		f.write(message)
		#20
		message = "y1_"+str(i)+" + y2_"+str(i)+" + y3_"+str(i)+" + y4_"+str(i)+" + y5_"+str(i)+" + y6_"+str(i)+" + y7_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,20)))+";\n"
		f.write(message)
		#21
		message = "y1_"+str(i)+" + y2_"+str(i)+" + y3_"+str(i)+" + y4_"+str(i)+" + y5_"+str(i)+" + y6_"+str(i)+" + y7_"+str(i)+" + y8_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,21)))+";\n"
		f.write(message)
		#22
		message = "y2_"+str(i)+" + y3_"+str(i)+" + y4_"+str(i)+" + y5_"+str(i)+" + y6_"+str(i)+" + y7_"+str(i)+" + y8_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,22)))+";\n"
		f.write(message)
		#23
		message = "y1_"+str(i)+" + y2_"+str(i)+" + y4_"+str(i)+" + y5_"+str(i)+" + y6_"+str(i)+" + y7_"+str(i)+" + y8_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,23)))+";\n"
		f.write(message)
		#24
		message = "y1_"+str(i)+" + y4_"+str(i)+" + y5_"+str(i)+" + y6_"+str(i)+" + y7_"+str(i)+" + y8_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,24)))+";\n"
		f.write(message)
		#25
		message = "y1_"+str(i)+" + y2_"+str(i)+" + y3_"+str(i)+" + y6_"+str(i)+" + y7_"+str(i)+" + y8_"+str(i)+" + y9_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,25)))+";\n"
		f.write(message)
		#26
		message = "y1_"+str(i)+" + y2_"+str(i)+" + y3_"+str(i)+" + y7_"+str(i)+" + y8_"+str(i)+" + y9_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,26)))+";\n"
		f.write(message)
		#27
		message = "y1_"+str(i)+" + y2_"+str(i)+" + y3_"+str(i)+" + y4_"+str(i)+" + y5_"+str(i)+" + y8_"+str(i)+" + y9_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,27)))+";\n"
		f.write(message)
		#28
		message = "y1_"+str(i)+" + y2_"+str(i)+" + y3_"+str(i)+" + y4_"+str(i)+" + y5_"+str(i)+" + y6_"+str(i)+" + y8_"+str(i)+" + y9_"+str(i)+" + y10_"+str(i)+" + y11_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,28)))+";\n"
		f.write(message)
		#29
		message = "y1_"+str(i)+" + y2_"+str(i)+" + y3_"+str(i)+" + y4_"+str(i)+" + y5_"+str(i)+" + y6_"+str(i)+" + y7_"+str(i)+" + y9_"+str(i)+" + y10_"+str(i)+" + y11_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,29)))+";\n"
		f.write(message)
		#30
		message = "y1_"+str(i)+" + y2_"+str(i)+" + y3_"+str(i)+" + y4_"+str(i)+" + y5_"+str(i)+" + y6_"+str(i)+" + y7_"+str(i)+" + y9_"+str(i)+" + y10_"+str(i)+" + y11_"+str(i)+" + y12_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,30)))+";\n"
		f.write(message)
		#31
		message = "y2_"+str(i)+" + y3_"+str(i)+" + y4_"+str(i)+" + y5_"+str(i)+" + y6_"+str(i)+" + y7_"+str(i)+" + y8_"+str(i)+" + y9_"+str(i)+" + y10_"+str(i)+" + y11_"+str(i)+" + y12_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,31)))+";\n"
		f.write(message)
		#32
		message = "y2_"+str(i)+" + y3_"+str(i)+" + y4_"+str(i)+" + y5_"+str(i)+" + y6_"+str(i)+" + y7_"+str(i)+" + y8_"+str(i)+" + y9_"+str(i)+" + y10_"+str(i)+" + y11_"+str(i)+" + y12_"+str(i)+" + y13_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,32)))+";\n"
		f.write(message)
		#33
		message = "y3_"+str(i)+" + y4_"+str(i)+" + y5_"+str(i)+" + y6_"+str(i)+" + y7_"+str(i)+" + y8_"+str(i)+" + y9_"+str(i)+" + y10_"+str(i)+" + y11_"+str(i)+" + y12_"+str(i)+" + y13_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,33)))+";\n"
		f.write(message)
		#34
		message = "y3_"+str(i)+" + y4_"+str(i)+" + y5_"+str(i)+" + y6_"+str(i)+" + y7_"+str(i)+" + y8_"+str(i)+" + y9_"+str(i)+" + y10_"+str(i)+" + y11_"+str(i)+" + y12_"+str(i)+" + y13_"+str(i)+" + y14_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,34)))+";\n"
		f.write(message)
		#35
		message = "y5_"+str(i)+" + y6_"+str(i)+" + y7_"+str(i)+" + y8_"+str(i)+" + y11_"+str(i)+" + y12_"+str(i)+" + y13_"+str(i)+" + y14_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,35)))+";\n"
		f.write(message)
		#36
		message = "y6_"+str(i)+" + y7_"+str(i)+" + y8_"+str(i)+" + y9_"+str(i)+" + y10_"+str(i)+" + y13_"+str(i)+" + y14_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,36)))+";\n"
		f.write(message)
		#37
		message = "y7_"+str(i)+" + y8_"+str(i)+" + y9_"+str(i)+" + y10_"+str(i)+" + y11_"+str(i)+" + y12_"+str(i)+" + y13_"+str(i)+" + y14_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,37)))+";\n"
		f.write(message)
		#38
		message = "y8_"+str(i)+" + y9_"+str(i)+" + y10_"+str(i)+" + y11_"+str(i)+" + y12_"+str(i)+" + y13_"+str(i)+" + y14_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,38)))+";\n"
		f.write(message)
		#39
		message = "y9_"+str(i)+" + y10_"+str(i)+" + y11_"+str(i)+" + y12_"+str(i)+" + y13_"+str(i)+" + y14_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,39)))+";\n"
		f.write(message)
		#40
		message = "y9_"+str(i)+" + y10_"+str(i)+" + y11_"+str(i)+" + y12_"+str(i)+" + y14_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,40)))+";\n"
		f.write(message)
		#41
		message = "y9_"+str(i)+" + y10_"+str(i)+" + y11_"+str(i)+" + y12_"+str(i)+" + y13_"+str(i)+" + y14_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,41)))+";\n"
		f.write(message)
		#42
		message = "y10_"+str(i)+" + y11_"+str(i)+" + y12_"+str(i)+" + y13_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,42)))+";\n"
		f.write(message)
		#43
		message = "y10_"+str(i)+" + y11_"+str(i)+" + y12_"+str(i)+" + y13_"+str(i)+" + y14_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,43)))+";\n"
		f.write(message)
		#44
		message = "y10_"+str(i)+" + y11_"+str(i)+" + y12_"+str(i)+" + y13_"+str(i)+" + y14_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,44)))+";\n"
		f.write(message)
		#45
		message = "y12_"+str(i)+" + y13_"+str(i)+" + y14_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,45)))+";\n"
		f.write(message)
		#46
		message = "y12_"+str(i)+" + y13_"+str(i)+" + y14_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,46)))+";\n"
		f.write(message)
		#47
		message = "y13_"+str(i)+" + y14_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,47)))+";\n"
		f.write(message)
		#48
		message = "y13_"+str(i)+" + y14_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,48)))+";\n"
		f.write(message)

#write weekend constraints

for i in weekends:
	if i%7 == 0:
		#1
		message = "y14_"+str(i-1)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,1)))+";\n"
		f.write(message)
		#2
		message = "y14_"+str(i-1)+" + y27_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,2)))+";\n"
		f.write(message)
	elif (i == 1):
		#2
		message = "y27_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,2)))+";\n"
	else:
		#1
		message = "y26_"+str(i-1)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,1)))+";\n"
		f.write(message)
		#2
		message = "y26_"+str(i-1)+" + y27_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,2)))+";\n"
		f.write(message)
	#3
	message = "y27_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,3)))+";\n"
	f.write(message)
	#4
	message = "y27_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,4)))+";\n"
	f.write(message)
	#5
	message = "y27_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,5)))+";\n"
	f.write(message)
	#6
	message = "y27_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,6)))+";\n"
	f.write(message)
	#7
	message = "y27_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,7)))+";\n"
	f.write(message)
	#8
	message = "y27_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,8)))+";\n"
	f.write(message)
	#9
	message = "y27_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,9)))+";\n"
	f.write(message)
	#10
	message = "y27_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,10)))+";\n"
	f.write(message)
	#11
	message = "y27_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,11)))+";\n"
	f.write(message)
	#12
	message = "y27_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,12)))+";\n"
	f.write(message)
	#13
	message = "y27_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,13)))+";\n"
	f.write(message)
	#14
	message = "y27_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,14)))+";\n"
	f.write(message)
	#15
	message = "y27_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,15)))+";\n"
	f.write(message)
	#16
	message = "y27_"+str(i)+" + y16_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,16)))+";\n"
	f.write(message)
	#17
	message = "y27_"+str(i)+" + y16_"+str(i)+" + y17_"+str(i)+" + y18_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,17)))+";\n"
	f.write(message)
	#18
	message = "y16_"+str(i)+" + y17_"+str(i)+" + y18_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,18)))+";\n"
	f.write(message)
	#19
	message = "y16_"+str(i)+" + y17_"+str(i)+" + y18_"+str(i)+" + y19_"+str(i)+" + y20_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,19)))+";\n"
	f.write(message)
	#20
	message = "y16_"+str(i)+" + y17_"+str(i)+" + y18_"+str(i)+" + y19_"+str(i)+" + y20_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,20)))+";\n"
	f.write(message)
	#21
	message = "y16_"+str(i)+" + y17_"+str(i)+" + y18_"+str(i)+" + y19_"+str(i)+" + y20_"+str(i)+" + y21_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,21)))+";\n"
	f.write(message)
	#22
	message = "y16_"+str(i)+" + y17_"+str(i)+" + y18_"+str(i)+" + y19_"+str(i)+" + y20_"+str(i)+" + y21_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,22)))+";\n"
	f.write(message)
	#23
	message = "y16_"+str(i)+" + y17_"+str(i)+" + y18_"+str(i)+" + y19_"+str(i)+" + y20_"+str(i)+" + y21_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,23)))+";\n"
	f.write(message)
	#24
	message = "y17_"+str(i)+" + y18_"+str(i)+" + y19_"+str(i)+" + y20_"+str(i)+" + y21_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,24)))+";\n"
	f.write(message)
	#25
	message = "y16_"+str(i)+" + y18_"+str(i)+" + y20_"+str(i)+" + y21_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,25)))+";\n"
	f.write(message)
	#26
	message = "y16_"+str(i)+" + y17_"+str(i)+" + y19_"+str(i)+" + y21_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,26)))+";\n"
	f.write(message)
	#27
	message = "y16_"+str(i)+" + y17_"+str(i)+" + y18_"+str(i)+" + y19_"+str(i)+" + y20_"+str(i)+" + y21_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,27)))+";\n"
	f.write(message)
	#28
	message = "y16_"+str(i)+" + y17_"+str(i)+" + y18_"+str(i)+" + y19_"+str(i)+" + y20_"+str(i)+" + y21_"+str(i)+" + y22_"+str(i)+" + y23_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,28)))+";\n"
	f.write(message)
	#29
	message = "y16_"+str(i)+" + y17_"+str(i)+" + y18_"+str(i)+" + y19_"+str(i)+" + y20_"+str(i)+" + y22_"+str(i)+" + y23_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,29)))+";\n"
	f.write(message)
	#30
	message = "y16_"+str(i)+" + y17_"+str(i)+" + y18_"+str(i)+" + y19_"+str(i)+" + y20_"+str(i)+" + y21_"+str(i)+" + y22_"+str(i)+" + y23_"+str(i)+" + y24_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,30)))+";\n"
	f.write(message)
	#31
	message = "y16_"+str(i)+" + y17_"+str(i)+" + y18_"+str(i)+" + y19_"+str(i)+" + y20_"+str(i)+" + y21_"+str(i)+" + y22_"+str(i)+" + y23_"+str(i)+" + y24_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,31)))+";\n"
	f.write(message)
	#32
	message = "y16_"+str(i)+" + y17_"+str(i)+" + y18_"+str(i)+" + y19_"+str(i)+" + y20_"+str(i)+" + y21_"+str(i)+" + y22_"+str(i)+" + y23_"+str(i)+" + y24_"+str(i)+" + y25_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,32)))+";\n"
	f.write(message)
	#33
	message = "y17_"+str(i)+" + y18_"+str(i)+" + y19_"+str(i)+" + y20_"+str(i)+" + y21_"+str(i)+" + y22_"+str(i)+" + y23_"+str(i)+" + y24_"+str(i)+" + y25_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,33)))+";\n"
	f.write(message)
	#34
	message = "y19_"+str(i)+" + y20_"+str(i)+" + y21_"+str(i)+" + y22_"+str(i)+" + y23_"+str(i)+" + y24_"+str(i)+" + y25_"+str(i)+" + y26_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,34)))+";\n"
	f.write(message)
	#35
	message = "y19_"+str(i)+" + y20_"+str(i)+" + y21_"+str(i)+" + y23_"+str(i)+" + y24_"+str(i)+" + y25_"+str(i)+" + y26_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,35)))+";\n"
	f.write(message)
	#36
	message = "y21_"+str(i)+" + y22_"+str(i)+" + y25_"+str(i)+" + y26_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,36)))+";\n"
	f.write(message)
	#37
	message = "y21_"+str(i)+" + y22_"+str(i)+" + y23_"+str(i)+" + y24_"+str(i)+" + y25_"+str(i)+" + y26_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,37)))+";\n"
	f.write(message)
	#38
	message = "y22_"+str(i)+" + y23_"+str(i)+" + y24_"+str(i)+" + y25_"+str(i)+" + y26_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,38)))+";\n"
	f.write(message)
	#39
	message = "y22_"+str(i)+" + y23_"+str(i)+" + y24_"+str(i)+" + y25_"+str(i)+" + y26_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,39)))+";\n"
	f.write(message)
	#40
	message = "y22_"+str(i)+" + y23_"+str(i)+" + y24_"+str(i)+" + y26_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,40)))+";\n"
	f.write(message)
	#41
	message = "y22_"+str(i)+" + y23_"+str(i)+" + y24_"+str(i)+" + y25_"+str(i)+" + y26_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,41)))+";\n"
	f.write(message)
	#42
	message = "y22_"+str(i)+" + y23_"+str(i)+" + y24_"+str(i)+" + y25_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,42)))+";\n"
	f.write(message)
	#43
	message = "y22_"+str(i)+" + y23_"+str(i)+" + y24_"+str(i)+" + y25_"+str(i)+" + y26_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,43)))+";\n"
	f.write(message)
	#44
	message = "y22_"+str(i)+" + y23_"+str(i)+" + y24_"+str(i)+" + y25_"+str(i)+" + y26_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,44)))+";\n"
	f.write(message)
	#45
	message = "y24_"+str(i)+" + y25_"+str(i)+" + y26_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,45)))+";\n"
	f.write(message)
	#46
	message = "y24_"+str(i)+" + y25_"+str(i)+" + y26_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,46)))+";\n"
	f.write(message)
	#47
	message = "y25_"+str(i)+" + y26_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,47)))+";\n"
	f.write(message)
	#48
	message = "y25_"+str(i)+" + y26_"+str(i)+" >= "+str(math.ceil(coverage*sheet.cell_value(i,48)))+";\n"
	f.write(message)

#write the variables used

f.write("int ")

message = ""

for i in range(4):
	for j in range(2+7*i, 7+7*i):
		for k in range(1, 16):
			message = message+", y"+str(k)+"_"+str(j)
			
weekends = (1, 7, 8, 14, 15, 21, 22, 28)

for i in weekends:
	for k in range(16, 28):
		message = message+", y"+str(k)+"_"+str(i)
		
message = message+";\n"
message = message[2:]

f.write(message)
f.close()