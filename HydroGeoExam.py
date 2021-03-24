import random
import xlrd

book = xlrd.open_workbook('~/Desktop/出题.xlsx')
xzt = book.sheets()[0]
jdt = book.sheets()[1]
ztt = book.sheets()[2]
jst = book.sheets()[3]
Nxzt = 135
Cxzt = 35
Njdt = 31
Nztt = 8
Njst = 7

h = open('题目.txt','w')
g = open('答案.txt','w')

groups = [[0 for i in range(5)] for j in range(13)]
candidates = list(range(Nxzt))
flag = 1
while flag:
	flag = 0
	random.shuffle(candidates)
	case = candidates[0:Cxzt]
	for i in range(Nxzt):
		zj = int(xzt.cell(i,0).value)
		fz = int(xzt.cell(i,1).value)
		if fz != 1:
			count = 0
			step = 0
			for j in range(10):
				if int(xzt.cell(i+j,1).value) != fz:
					break
				else:
					step += 1
					idx = i+j
					if idx in case:
						count += 1
			if count > 1:
				flag = 1
				break
			else:
				i += step
h.write('一、选择题\n')
g.write('一、选择题\n')
for i in range(Cxzt):
	idx = case[i]
	timu = xzt.cell(idx,2).value
	A = xzt.cell(idx,3).value
	B = xzt.cell(idx,4).value
	C = xzt.cell(idx,5).value
	D = xzt.cell(idx,6).value
	ans = xzt.cell(idx,7).value
	h.write(str(i+1)+'. '+str(timu)+'\n'+'A '+str(A)+'B '+str(B)+'C '+str(C)+'D '+str(D)+'\n')
	g.write(ans)
	if i%5 == 4:
		g.write('  ')
g.write('\n')
######################################
h.write('二、绘图题\n')
g.write('二、绘图题\n')
candidates = list(range(6))
random.shuffle(candidates)
case1 = candidates[0:1]
h.write(str(1)+'. 请绘制下图所示的渗流场内的地下水流网图 (插入Fig)'+str(case1[0])+'\n')
candidates = list(range(2))
random.shuffle(candidates)
case1 = candidates[0:1]
idx = case1[0]+1
timu = ztt.cell(idx,2).value
h.write(str(2)+'. '+str(timu)+'\n')
candidates = list(range(2))
random.shuffle(candidates)
case1 = candidates[0:1]
idx = case1[0]+3
timu = ztt.cell(idx,2).value
h.write(str(3)+'. '+str(timu)+' (插入Fig)'+str(case1[0])+'\n')
candidates = list(range(3))
random.shuffle(candidates)
case1 = candidates[0:1]
idx = case1[0]+5
timu = ztt.cell(idx,2).value
h.write(str(4)+'. '+str(timu)+'\n')
#######################################
h.write('三、简答题\n')
g.write('三、简答题\n')
candidates = list(range(20))
random.shuffle(candidates)
case1 = candidates[0:4]
candidates = list(range(10))
random.shuffle(candidates)
case2 = candidates[0:2]
for i in range(4):
	idx = case1[i]
	timu = jdt.cell(idx,2).value
	ans = jdt.cell(idx,3).value
	h.write(str(i+1)+'. '+str(timu)+'\n')
	g.write(str(i+1)+'. '+str(ans)+'\n')
for i in range(2):
	idx = case2[i] + 20
	timu = jdt.cell(idx,2).value
	ans = jdt.cell(idx,3).value
	h.write(str(i+5)+'. '+str(timu)+'\n')
	g.write(str(i+5)+'. '+str(ans)+'\n')
#######################################
h.write('四、计算题\n')
g.write('四、计算题\n')
candidates = list(range(7))
random.shuffle(candidates)
case1 = candidates[0:1]
idx = case1[0]
timu = jst.cell(idx,1).value
ans = jst.cell(idx,2).value
h.write(str(1)+'. '+str(timu)+' (插入Fig)'+'\n')
g.write(str(1)+'. '+str(ans)+' (插入Fig)'+'\n')
#######################################
h.write('五、写作题\n')
h.write('1. 请谈一谈你觉得重庆地区的地下水的赋存、补给、径流、排泄、水化学特征等的特点（200字以内）。\n')
#######################################
g.close()
h.close()
