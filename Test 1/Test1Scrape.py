import os

cwd = os.getcwd()
cwd

os.chdir("C:/Users/Reube/Desktop/New Folder")

os.listdir('.')

import pandas as pd
import numpy as np

file = 'Test1.xls'


x1 = pd.ExcelFile(file)

#print(x1.sheet_names)
df1 = x1.parse('Sheet1')

import xlrd
x1_workbook = xlrd.open_workbook(file)
# from openpyxl import load_workbook
#
# wb= load_workbook('Test1.xls')
# print(wb.get_sheet_names())
# print (x1_workbook.nsheets)
# print (x1_workbook.sheet_names())
first_sheet = x1_workbook.sheet_by_index(0)
headers = (first_sheet.row_values(2))
headers.remove("")

df = pd.DataFrame(columns=headers)


for x in range(4,21):
    data = {
    'DATE':xlrd.xldate_as_datetime(int(first_sheet.cell(x,0).value), x1_workbook.datemode),
    'FLEET NO.': first_sheet.cell(x,1).value,
    'LOCATION': first_sheet.cell(x,3).value,
    'DIRECTION': first_sheet.cell(x,4).value,
    'ACCIDENT DETAILS': first_sheet.cell(x,5).value
    }
    df = df.append(data,ignore_index=True)


print (df)

df.to_csv("export.csv")
