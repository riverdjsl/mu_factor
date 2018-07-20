import numpy as np
import xlwings as xw

wb = xw.Book('newmodel.xlsx')
OSC = wb.sheets['Over - Steel - Chinese 2010']


def trans_comma_str(string):
	'''to transform comma strings into lists of numbers'''
	item = []
	lst = []
	for i in string[:-1]:
		if i != ",":
			item.append(i)
		else:
			lst.append(float(''.join(item)))
			item = []
			continue
	return lst


#loc = "I" or "J"	
def writetable(datafile, loc):
	data = {}
	with open(datafile) as f1:
		for i in f1.readlines():
			a = trans_comma_str(i)
			data[str(int(a[0]))] = a[1]
		f1.close()
	
	b = set(data.keys())

	k = 4
	while True:
		c = OSC.range("A"+str(k)).value
		if c in b:
			OSC.range(loc+str(k)).value = data[c]
		elif c is not None:
			OSC.range(loc+str(k)).value = 0
		else:
			break
		k += 1
		



writetable("out.txt", "I")
writetable("outh.txt", "J")		
		

	
