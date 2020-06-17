import xlsxwriter
import xlrd

#reorganize people by group

location1 = ("peoplelabel_original.xlsx")

workbook1 = xlrd.open_workbook(location1)
worksheet1 = workbook1.sheet_by_index(0) 

workbook2 = xlsxwriter.Workbook('peoplelabel_organized.xlsx')
worksheet2 = workbook2.add_worksheet()

numberofpeople = worksheet1.nrows-1

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

for i in range(1, worksheet1.nrows):
	if (worksheet1.cell_value(i,1) == "高雄" and worksheet1.cell_value(i,2) == "F"):
		KF = KF + [int(worksheet1.cell_value(i,0))]
	elif(worksheet1.cell_value(i,1) == "台北" and worksheet1.cell_value(i,2) == "F"):
		TF = TF + [int(worksheet1.cell_value(i,0))]
	elif(worksheet1.cell_value(i,1) == "台中" and worksheet1.cell_value(i,2) == "F"):
		CF = CF + [int(worksheet1.cell_value(i,0))]
	elif(worksheet1.cell_value(i,1) == "高雄" and worksheet1.cell_value(i,2) == "M"):
		KM = KM + [int(worksheet1.cell_value(i,0))]
	elif(worksheet1.cell_value(i,1) == "台北" and worksheet1.cell_value(i,2) == "M"):
		TM = TM + [int(worksheet1.cell_value(i,0))]
	elif(worksheet1.cell_value(i,1) == "台中" and worksheet1.cell_value(i,2) == "M"):
		CM = CM + [int(worksheet1.cell_value(i,0))]

F = F+KF+TF+CF
M = M+KM+TM+CM
K = K+KF+KM
T = T+TF+TM
C = C+CF+CM
All = All+F+M

K1 = []
K2 = []
K3 = []
K4 = []
K5 = []
K6 = []
K7 = []

for i in KF:
	if (worksheet1.cell_value(i,3) == "K1"):
		K1 = K1 + [worksheet1.cell_value(i,0)]	
	elif (worksheet1.cell_value(i,3) == "K2"):
		K2 = K2 + [worksheet1.cell_value(i,0)]	
	elif (worksheet1.cell_value(i,3) == "K3"):
		K3 = K3 + [worksheet1.cell_value(i,0)]	
	elif (worksheet1.cell_value(i,3) == "K4"):
		K4 = K4 + [worksheet1.cell_value(i,0)]	
	elif (worksheet1.cell_value(i,3) == "K5"):
		K5 = K5 + [worksheet1.cell_value(i,0)]	
	elif (worksheet1.cell_value(i,3) == "K6"):
		K6 = K6 + [worksheet1.cell_value(i,0)]	
	elif (worksheet1.cell_value(i,3) == "K7"):
		K7 = K7 + [worksheet1.cell_value(i,0)]	

worksheet2.write(0, 0, worksheet1.cell_value(0,0))
worksheet2.write(0, 1, worksheet1.cell_value(0,1))
worksheet2.write(0, 2, worksheet1.cell_value(0,2))
worksheet2.write(0, 3, worksheet1.cell_value(0,3))

counter = 1

for i in range(len(K1)):
	worksheet2.write(counter, 0, K1[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(K1[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(K1[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(K1[i]),3))
	counter = counter+1

for i in range(len(K2)):
	worksheet2.write(counter, 0, K2[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(K2[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(K2[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(K2[i]),3))
	counter = counter+1

for i in range(len(K3)):
	worksheet2.write(counter, 0, K3[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(K3[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(K3[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(K3[i]),3))
	counter = counter+1

for i in range(len(K4)):
	worksheet2.write(counter, 0, K4[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(K4[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(K4[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(K4[i]),3))
	counter = counter+1

for i in range(len(K5)):
	worksheet2.write(counter, 0, K5[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(K5[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(K5[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(K5[i]),3))
	counter = counter+1

for i in range(len(K6)):
	worksheet2.write(counter, 0, K6[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(K6[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(K6[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(K6[i]),3))
	counter = counter+1

for i in range(len(K7)):
	worksheet2.write(counter, 0, K7[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(K7[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(K7[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(K7[i]),3))
	counter = counter+1
	
T1 = []
T2 = []
T3 = []	

for i in TF:
	if (worksheet1.cell_value(i,3) == "T1"):
		T1 = T1 + [worksheet1.cell_value(i,0)]	
	elif (worksheet1.cell_value(i,3) == "T2"):
		T2 = T2 + [worksheet1.cell_value(i,0)]	
	elif (worksheet1.cell_value(i,3) == "T3"):
		T3 = T3 + [worksheet1.cell_value(i,0)]

for i in range(len(T1)):
	worksheet2.write(counter, 0, T1[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(T1[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(T1[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(T1[i]),3))
	counter = counter+1

for i in range(len(T2)):
	worksheet2.write(counter, 0, T2[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(T2[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(T2[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(T2[i]),3))
	counter = counter+1

for i in range(len(T3)):
	worksheet2.write(counter, 0, T3[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(T3[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(T3[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(T3[i]),3))
	counter = counter+1
	
C1 = []
C2 = []	

for i in CF:
	if (worksheet1.cell_value(i,3) == "C1"):
		C1 = C1 + [worksheet1.cell_value(i,0)]	
	elif (worksheet1.cell_value(i,3) == "C2"):
		C2 = C2 + [worksheet1.cell_value(i,0)]	

for i in range(len(C1)):
	worksheet2.write(counter, 0, C1[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(C1[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(C1[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(C1[i]),3))
	counter = counter+1

for i in range(len(C2)):
	worksheet2.write(counter, 0, C2[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(C2[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(C2[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(C2[i]),3))
	counter = counter+1
	
K1 = []
K2 = []
K3 = []
K4 = []
K5 = []
K6 = []
K7 = []

for i in KM:
	if (worksheet1.cell_value(i,3) == "K1"):
		K1 = K1 + [worksheet1.cell_value(i,0)]	
	elif (worksheet1.cell_value(i,3) == "K2"):
		K2 = K2 + [worksheet1.cell_value(i,0)]	
	elif (worksheet1.cell_value(i,3) == "K3"):
		K3 = K3 + [worksheet1.cell_value(i,0)]	
	elif (worksheet1.cell_value(i,3) == "K4"):
		K4 = K4 + [worksheet1.cell_value(i,0)]	
	elif (worksheet1.cell_value(i,3) == "K5"):
		K5 = K5 + [worksheet1.cell_value(i,0)]	
	elif (worksheet1.cell_value(i,3) == "K6"):
		K6 = K6 + [worksheet1.cell_value(i,0)]	
	elif (worksheet1.cell_value(i,3) == "K7"):
		K7 = K7 + [worksheet1.cell_value(i,0)]	

for i in range(len(K1)):
	worksheet2.write(counter, 0, K1[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(K1[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(K1[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(K1[i]),3))
	counter = counter+1

for i in range(len(K2)):
	worksheet2.write(counter, 0, K2[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(K2[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(K2[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(K2[i]),3))
	counter = counter+1

for i in range(len(K3)):
	worksheet2.write(counter, 0, K3[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(K3[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(K3[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(K3[i]),3))
	counter = counter+1

for i in range(len(K4)):
	worksheet2.write(counter, 0, K4[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(K4[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(K4[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(K4[i]),3))
	counter = counter+1

for i in range(len(K5)):
	worksheet2.write(counter, 0, K5[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(K5[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(K5[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(K5[i]),3))
	counter = counter+1

for i in range(len(K6)):
	worksheet2.write(counter, 0, K6[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(K6[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(K6[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(K6[i]),3))
	counter = counter+1

for i in range(len(K7)):
	worksheet2.write(counter, 0, K7[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(K7[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(K7[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(K7[i]),3))
	counter = counter+1

T1 = []
T2 = []
T3 = []	

for i in TM:
	if (worksheet1.cell_value(i,3) == "T1"):
		T1 = T1 + [worksheet1.cell_value(i,0)]	
	elif (worksheet1.cell_value(i,3) == "T2"):
		T2 = T2 + [worksheet1.cell_value(i,0)]	
	elif (worksheet1.cell_value(i,3) == "T3"):
		T3 = T3 + [worksheet1.cell_value(i,0)]

for i in range(len(T1)):
	worksheet2.write(counter, 0, T1[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(T1[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(T1[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(T1[i]),3))
	counter = counter+1

for i in range(len(T2)):
	worksheet2.write(counter, 0, T2[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(T2[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(T2[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(T2[i]),3))
	counter = counter+1

for i in range(len(T3)):
	worksheet2.write(counter, 0, T3[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(T3[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(T3[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(T3[i]),3))
	counter = counter+1

C1 = []
C2 = []	

for i in CM:
	if (worksheet1.cell_value(i,3) == "C1"):
		C1 = C1 + [worksheet1.cell_value(i,0)]	
	elif (worksheet1.cell_value(i,3) == "C2"):
		C2 = C2 + [worksheet1.cell_value(i,0)]	

for i in range(len(C1)):
	worksheet2.write(counter, 0, C1[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(C1[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(C1[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(C1[i]),3))
	counter = counter+1

for i in range(len(C2)):
	worksheet2.write(counter, 0, C2[i])
	worksheet2.write(counter, 1, worksheet1.cell_value(int(C2[i]),1))
	worksheet2.write(counter, 2, worksheet1.cell_value(int(C2[i]),2))
	worksheet2.write(counter, 3, worksheet1.cell_value(int(C2[i]),3))
	counter = counter+1
	
workbook2.close()