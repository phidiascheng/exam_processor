import random
import xlrd

book = xlrd.open_workbook('~/Desktop/专业英语题库.xlsx')
sheet = book.sheets()[0]
sheetfyt = book.sheets()[1]
sheetwdt = book.sheets()[2]
Nxzt = 131
Nfyt = 19
Nwdt = 3
Exzt = 25
Efyt = 6
Ewdt = 1
xzt = list(range(Nxzt))
fyt = list(range(Nfyt))
wdt = list(range(Nwdt))

f = open("专业英语试题.txt","w")
g = open('all.txt','w')
h = open('答案.txt','w')
DanCi = []
FanYi = []

for p in range(3):
	random.shuffle(xzt)
	random.shuffle(fyt)
	random.shuffle(wdt)
	f.write('试题'+str(p+1)+'\n\n')
	f.write('一、请翻译以下英文单词:'+'\n')
	h.write('试题'+str(p+1)+'\n\n')
	h.write('一、请翻译以下英文单词:'+'\n')
	for i in range(Exzt):
		TiMu = sheet.cell(xzt[i],0).value
		if TiMu not in DanCi:
			DanCi.append(TiMu)
		X = list(range(Nxzt))
		random.shuffle(X)
		XX = X[0:4]
		if xzt[i] not in XX:
			XX[3] = xzt[i]
		random.shuffle(XX)
		f.write(str(i+1)+'. '+str(TiMu)+'\n')
		f.write('A '+str(sheet.cell(XX[0],1).value))
		f.write(' B '+str(sheet.cell(XX[1],1).value))
		f.write(' C '+str(sheet.cell(XX[2],1).value))
		f.write(' D '+str(sheet.cell(XX[3],1).value)+'\n')
		h.write(str(i+1)+'. ')
		ans = XX.index(xzt[i])
		if ans == 0:
			h.write('A ')
		elif ans == 1:
			h.write('B ')
		elif ans == 2:
			h.write('C ')
		else:
			h.write('D ')

	f.write('\n'+'二、请翻译以下中文单词:'+'\n')
	h.write('\n'+'二、请翻译以下中文单词:'+'\n')
	for j in range(Exzt):
		i = j + Exzt
		TiMu = sheet.cell(xzt[i],0).value
		if TiMu not in DanCi:
			DanCi.append(TiMu)
		TiMu = sheet.cell(xzt[i],1).value
		X = list(range(Nxzt))
		random.shuffle(X)
		XX = X[0:4]
		if xzt[i] not in XX:
			XX[3] = xzt[i]
		random.shuffle(XX)
		f.write(str(j+1)+'. '+str(TiMu)+'\n')
		f.write('A '+str(sheet.cell(XX[0],0).value))
		f.write(' B '+str(sheet.cell(XX[1],0).value))
		f.write(' C '+str(sheet.cell(XX[2],0).value))
		f.write(' D '+str(sheet.cell(XX[3],0).value)+'\n')
		h.write(str(j+1)+'. ')
		ans = XX.index(xzt[i])
		if ans == 0:
			h.write('A ')
		elif ans == 1:
			h.write('B ')
		elif ans == 2:
			h.write('C ')
		else:
			h.write('D ')

	f.write('\n'+'三、请翻译以下英文段落:'+'\n')
	h.write('\n'+'三、请翻译以下英文段落:'+'\n')
	for i in range(Efyt):
		TiMu = sheetfyt.cell(fyt[i],0).value
		if TiMu not in FanYi:
			FanYi.append(TiMu)
		f.write(str(i+1)+'. '+str(TiMu)+'\n\n')
		h.write(str(i+1)+'. '+sheetfyt.cell(fyt[i],1).value+'\n\n')

	f.write('\n'+'四、问答题:'+'\n')
	h.write('\n'+'四、问答题:'+'\n')
	f.write('1. Please write a story of any scientist whose name appears in your college textbook. (请任意选择一位在你的大学教材上出现过名字的科学家，然后用英文介绍关于这位科学家的故事)')
	h.write('1. Please write a story of any scientist whose name appears in your college textbook. (请任意选择一位在你的大学教材上出现过名字的科学家，然后用英文介绍关于这位科学家的故事)')
	f.write('\n\n')
	h.write('\n\n')
	for i in range(Ewdt):
		TiMu = sheetwdt.cell(wdt[i],0).value
		f.write(str(i+2)+'. '+str(TiMu)+'\n\n')
		h.write(str(i+2)+'. '+str(TiMu)+'\n\n')

g.write('单词汇总:\n\n')
for i in range(len(DanCi)):
	g.write(str(i+1)+'.  '+str(DanCi[i])+'\n\n')
g.write('翻译汇总:\n\n')
for i in range(len(FanYi)):
	g.write(str(i+1)+'.  '+str(FanYi[i])+'\n\n')
f.close()
g.close()