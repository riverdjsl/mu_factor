import numpy as np
import xlwings as xw

wb1 = xw.Book('buckling.xlsx')
bfs1 = wb1.sheets['Buckling Factors']
bf1 = bfs1.range('D4').expand().value[0:20]

wb2 = xw.Book('bucklingh.xlsx')
bfs2 = wb2.sheets['Buckling Factors']
bf2 = bfs2.range('D4').expand().value[20:]


def mu(bf, I, p, l, E=206e6):
	results = []
	for i in bf:
		if i*p < 0:
			results.append(np.pi*np.sqrt(E*I/(-i*p))/l)
		else:
			results.append(0)
	return max(results)


def read_single_column(table,location):
	i = int(location[1:])
	c = location[0]
	while True:
		value = table.range(c+str(i)).value
		if value is not None:
			yield value
		else:
			break
		i += 1


def read_double_column(table, location1, location2):
	i1 = int(location1[1:])
	c1 = location1[0]
	i2 = int(location2[1:])
	c2 = location2[0]
	while True:
		value = [table.range(c1+str(i1)).value, table.range(c2+str(i2)).value]
		if value[0] is not None:
			yield value
		else:
			break
		i1 += 1
		i2 += 1


def read_triple_column(table, location1, location2, location3):
	i1 = int(location1[1:])
	c1 = location1[0]
	i2 = int(location2[1:])
	c2 = location2[0]
	i3 = int(location3[1:])
	c3 = location3[0]
	while True:
		value = [table.range(c1+str(i1)).value, table.range(c2+str(i2)).value, table.range(c3+str(i3)).value]
		if value[0] is not None:
			yield value
		else:
			break
		i1 += 1
		i2 += 1
		i3 += 1


def fuckit(xls, bf, ixx, fout):
	efs = xls.sheets['Element Forces - Frames']
	cfs = xls.sheets['Connectivity - Frame']
	fsas = xls.sheets['Frame Section Assignments']
	fps = xls.sheets['Frame Props 01 - General']
	dic = {}
	ppt = fps.range('A4').expand().value

	for i in read_single_column(cfs, 'A4'):
		dic[i] = []

	for i in read_double_column(fsas, 'A4', 'D4'):
		for j in ppt:
			if i[1] == j[0]:
				try:
					dic[i[0]].append(j[ixx])
				except KeyError:
					print('lala')

	for i in read_triple_column(efs, 'A4', 'E4', 'B4'):
		if i[2] == 0:
			dic[i[0]].append(i[1])

	for i in read_double_column(cfs, 'A4', 'E4'):
		dic[i[0]].append(i[1])

	f = open(fout, 'w')
	for i, j in dic.items():
		f.write("{},{:1f},".format(i, mu(bf, *j)))
		f.write('\n')
	f.close()

fuckit(wb1, bf1, 11, 'out.txt')
fuckit(wb2, bf2, 12, 'outh.txt')
