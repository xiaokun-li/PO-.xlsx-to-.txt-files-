#  Test some expected function about Excel documents.

import math
import xlrd
import os
import re

class MyFirstClass:

    def __init__(self, x=0, y=0):
        self.move(x,y)
    
    def move(self,x,y):
        self.x = x
        self.y = y

    def reset(self):
        self.move(0,0)

    def caculate_distance(self, other_point):

        assert (self.x - other_point.x)**2 + (self.y - other_point.y)**2 == (other_point.x-self.x)**2 + (other_point.y-self.y )**2
        return math.sqrt((self.x - other_point.x)**2 + (self.y - other_point.y)**2)


targetfile = r'KS - CPU.xlsx'


a = MyFirstClass()
b = MyFirstClass()

a.move(1,2)
b.move(6,6)



print(a,b)
print(a.x, a.y,b.x,b.y)
print(a.caculate_distance(b))


def is_number(num):
    pattern = re.compile(r'^[-+]?[-0-9]\d*\.\d*|[-+]?\.?[0-9]\d*$')
    result = pattern.match(num)
    if result:
        return True
    else:
        return False

excelfilelist = []

def getExcelfiles():
    for i in os.listdir():
        if i.endswith('.xlsx'):
            excelfilelist.append(i)


def poxls2txt(excelfile):
    tworkbook = xlrd.open_workbook(excelfile)
    tworksheet = tworkbook.sheets()[0]
    rowcount = tworksheet.nrows
    rowlist = []
    rowlist2 = []

    if rowcount>=2:
        for row in range(1,rowcount):
            rowlist.append(tworksheet.row_values(row)[:11])
    

    for row in rowlist:
        row0 = '' if row[0]=='' else (str(row[0]).partition('.'))[0]
        row2 = '' if row[2]=='' else (str(row[2]).partition('.'))[0]
        row5 = '' if row[5]=='' else (str(int(row[5])) if is_number(str(row[5])) else str(row[5]))
        row10 = '' if row[10]=='' else str(row[10])
        iss =  row0+'\t'+str(row[1])+'\t' + row2 + '\t' +str(row[3]) + '\t' + str(row[4]) + '\t' + row5 + '\t' + str(row[6]) + '\t' + str(row[7]) + '\t' + str(row[8]) + '\t' + str(row[9]) + '\t' + row10 +'\n'
        rowlist2.append(iss)

    f = open(excelfile.partition(".xls")[0]+'20200703.txt','w')
    for i in rowlist2:
        f.write(i)
    f.close()



getExcelfiles()

for i in (excelfilelist):
    poxls2txt(i)





#  ii2 = '' if i[2]=='' else i[2]
#  ii10 = '' if i[10]=='' else i[10]

#  str(int(i[0]))+'\t'+i[1]+'\t' + ii2 + '\t' +i[3] + '\t' + i[4] + '\t' + i[5] + '\t' + i[6] + '\t' + str(i[7]) + '\t' + str(i[8]) + '\t' + str(i[9]) + '\t' + ii10 



#  test code:

