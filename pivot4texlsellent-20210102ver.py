#pivot4texlsellent-20210102ver.py 
#20201231-20210102, DRAFT?

# coding:utf-8
#pip3 install xlrd==1.2.0

import time
import pprint
import openpyxl
import xlrd
import xlwt
from openpyxl import load_workbook

import pandas as pd

#Sample (of Texlsellent)をDataFrameとして取得しているらしい
df1 = pd.read_excel('Sample4pivot.xlsx', sheet_name = 'Texls1')

#試しに集計してみる
piv1= df1.pivot_table(index=['項番','aa','sum'], values=['viva','textjoin'],aggfunc='sum')

piv2= df1.pivot_table(index=['aa'], values=['viva','textjoin','dd'],aggfunc='sum')

piv3= df1.pivot_table(index=['aa'], values=['mm','viva','textjoin'],aggfunc='sum')

print(piv1)
print('------')
print(piv2)
print('------')
print(piv3)

piv1.to_excel('Sample4pivot2.xlsx', sheet_name='pivotexls2')

piv2.to_excel('Sample4pivot3.xlsx', sheet_name='pivotexls3')

piv3.to_excel('Sample4pivot4.xlsx', sheet_name='pivotexls4')

#xlrd1.2.0だと、できた20210101


# “pivot4texlsellent in Python 3.8” tested by K.masamix “KIXAN” as the SHYNAMITES/La CHENAMITOJ, 2020/12/31-2021/01/02.