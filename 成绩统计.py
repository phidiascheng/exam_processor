import xlrd, numpy, datetime, time
import matplotlib
import matplotlib.pyplot as plt 
import matplotlib.lines as mlines
from matplotlib import dates
from brokenaxes import brokenaxes

book = xlrd.open_workbook('~/Desktop/45的副本.xlsx')
sheet = book.sheets()[0]


N = 44

xuanze = [['' for i in range(50)] for j in range(N)]
fanyi = [[0 for i in range(8)] for j in range(N)]
name = ['' for i in range(N)]
Daan = ['B','C','B','C','B','C','B','B','A','D','A','D','C','D','D','C','D','A','B','A','B','A','D','C','A',
		'C','C','D','D','A','D','C','C','A','D','A','B','C','D','A','B','A','A','D','C','D','B','D','B','C']

for p in range(N):
	name[p] = sheet.cell(p,2).value
	for i in range(8):
		fanyi[p][i] = sheet.cell(p,14+i).value
	for i in range(10):
		temp = list(sheet.cell(p,4+i).value)
		for j in range(5):
			if temp == []:
				xuanze[p][i*5+j] = 'X'
			else:
				xuanze[p][i*5+j] = temp[j]

data=open("45.txt",'w')
for p in range(N):
	score1 = 0
	score2 = 0
	score3 = 0
	score4 = 0
	print(name[p],'\t',end='')
	data.write(name[p]+'\t')
	for i in range(50):
		if xuanze[p][i] != Daan[i]:
			print("\033[0;37m%s\033[0m"%xuanze[p][i],end='' )
			score1 += 1
			if i<25:
				score4 += 1
		else:
			print("\033[0;33m%s\033[0m"%xuanze[p][i],end='' )
		if i%5 == 4:
			print(' ',end='')
	print('  ',end='')
	for i in range(8):
		if fanyi[p][i] == '':
			print('0 ',end='')
		else:
			print(int(fanyi[p][i]),' ',end='')
			score2 += int(fanyi[p][i])
			if i<6:
				score3 += int(fanyi[p][i])
	print('总分',25-score4, 50-score1-25+score4, score3, score2-score3, 50-score1+score2)
	data.write(str(25-score4)+'\t'+str(25-score1+score4)+'\t'+str(50-score1+score2)+'\n')
data.close()
