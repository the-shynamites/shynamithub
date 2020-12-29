#HyperLinxionary-20201229ver.py
# coding:utf-8

import time
import openpyxl
from openpyxl import load_workbook

# 単語リストを含むサンプルのワークブック「HyperLinxionary0.xlsx」を開く(Excelファイルを読み込む)。
book0 = openpyxl.load_workbook('HyperLinxionary0.xlsx')

#シート「Linxionary」を指定する
sheet0 = book0['Linxionary']

#【例1】辞書サイトA(URL→https://仮frjp.jpnn)で検索した場合。

cellsC = sheet0['C2':'C30']
for row2 in cellsC:
    for cell2 in row2:
         cell2.value = '=HYPERLINK(TEXTJOIN("",TRUE,"https://frjp.jpnn/",OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())), 0, -1)),OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())), 0, -1))'


# 1秒間の待機
time.sleep(1)

#【例2】辞書サイトB(URL→https://仮frfrjpjp)で検索した場合。

cellsD = sheet0['D2':'D30']
for row3 in cellsD:
    for cell3 in row3:
         cell3.value = '=HYPERLINK(TEXTJOIN("",TRUE,"https://仮frfrjpjp-",OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())), 0, -2),".html"),OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())), 0, -2))' 

# 1秒間の待機
time.sleep(1)


#元のワークブックに上書き保存する
book0.save('HyperLinxionary0.xlsx')


# “HyperLinxionary (20201229ver.) in Python 3.8” produced by K.masamix “KIXAN” as the SHYNAMITES/La CHENAMITOJ, 2020/12/29.

