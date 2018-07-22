import numpy as np
import xlwings as xw
import time

bucfacbook = xw.Book('buckfac.xlsx')
bfs = bucfacbook.sheets['Buckling Factors']

def GetNowTime():
    return time.strftime("%Y-%m-%d %H:%M:%S",time.localtime(time.time()))


def mui(I, p, l, E=206e6):
	return (np.pi*np.sqrt(E*I/(-p))/l)

#rxx 11 for i33, 12 for i22
def mu(book, a, b, rxx):
	wb = xw.Book(book)
	bf = bfs.range('D4').expand().value[a:b]
	fsas = wb.sheets['Frame Section Assignments']
	cfs = wb.sheets['Connectivity - Frame']

	log = open('log.txt', 'a+')
	log.write('\n')
	log.write('\n')

	title0 = 'Starting processing {} at {}'.format(book, GetNowTime())
	log.write(title0)
	print(title0)
	log.write('\n')

	members = {}
	for i in cfs.range('A4').expand().value:
		members[i[0]] = i[4]

	log.write('The total number of members is {}.'.format(len(list(members.keys()))))

	sec = fsas.range('D4').value
	for i in wb.sheets['Frame Props 01 - General'].range('A4').expand().value:
		if i[0] == sec:
			ixx = i[rxx]

	log.write('\n')
	log.write('The section is {}.'.format(sec))
	log.write('\n')
	log.write('The 2nd inertia moment is {}.'.format(ixx))
	log.write('\n')

	def disp(sht, n):
		k = 0
		while True:
			key = str(4+n+k*len(bf))
			try:
				a = sht.range('A'+key).value
				d1 = sht.range('F'+key).value
				d2 = sht.range('G'+key).value
				d3 = sht.range('H'+key).value
				disp = np.sqrt(d1**2+d2**2+d3**2)
				k += 1
				yield (a, disp)
			except Exception:
				return

	def maxnode(k):
		dic = {}
		disps = []
		for i in disp(wb.sheets['Joint Displacements - Absolute'], k):
			dic[i[0]] = i[1]
			disps.append(i[1])
		maxdisp = max(disps)
		return list(dic.keys())[list(dic.values()).index(maxdisp)], maxdisp

	def frame(node):
		for i in cfs.range('A4').expand().value:
			if i[2] == node:
				return i[0], i[4]

	log.write('\n')
	title1 = '{:>10}{:>10}{:>10}{:>10}{:>10}'.format('Mode', 'Factor', 'Node', 'Disp', 'Frame')
	log.write(title1)
	print(title1)
	log.write('\n')

	data = {}
	frame_critical = []
	for i,j in enumerate(bf):
		nodei = maxnode(i)[0]
		maxnodedispi = maxnode(i)[1]
		framei = frame(nodei)
		data[framei[0]] = [i, j, framei[1]]
		frame_critical.append(framei[0])

		content1 = '{:10}{:10.2f}{:>10}{:10.2e}{:>10}'.format(i, j, nodei, maxnodedispi, framei[0])
		log.write(content1)
		print(content1)
		log.write('\n')

	frame_critical = set(frame_critical)

	efs = wb.sheets['Element Forces - Frames']
	for i, j in zip(efs.range('A4').expand().value, efs.range('G4').expand().value):
		if i[0] in frame_critical:
			if i[1] == 0:
				data[i[0]].append(j[0])

	log.write('\n')
	title2 = '{:>10}{:>10}{:>10}{:>10}{:>10}'.format('Frame', 'Mode', 'Factor', 'Pcr', 'Mu')
	log.write(title2)
	print(title2)
	log.write('\n')

	mus = []
	for j, i in data.items():
		p = i[-len(bf):][i[0]]
		if p/i[1] < 0 and p < 0:
			muii = mui(ixx, p, i[2])
			mus.append([muii, i[2]])

			content2 = '{:>10}{:10}{:10.2f}{:10.2f}{:10.2f}'.format(j, i[0], i[1], p, muii)
			log.write(content2)
			print(content2)
			log.write('\n')

	question = 'Now choose a mu according to your judgement!:'

	signal = True
	while signal:
		try:
			keymu = input(question)
			for i, j in enumerate(mus):
				if i == int(keymu):
					l0 = j[0]*j[1]
					signal = False
			else:
				print('Out of range, try again!')
		except Exception:
			print('Number!Please!!')

	log.write('\n')
	log.write('{} {}'.format(question, keymu))
	log.write('\n')

	muall = {}
	for i,j in members.items():
		muall[i] = l0/j

	outfilename = 'out_'+str(rxx)+'_'+book[:-5]+'.txt'
	outfile = open(outfilename, 'w')
	for i, j in muall.items():
		outfile.write('{},{:5.2f},'.format(i, j))
		outfile.write('\n')
	outfile.close()

	log.write('\n')
	log.write('Find the results for {} in {}.'.format(book,outfilename))

	log.close()
	print("done for now")

	return muall



with open('log.txt', 'w') as f:
	f.close()

print(mu('buckpattern2_upperring.xlsx', 20, 40, 12))
print(mu('buckpattern1_lowerring.xlsx', 0, 20, 11))
print(mu('buckpattern1_upperring.xlsx', 0, 20, 11))

