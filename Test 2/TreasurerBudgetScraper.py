import pandas as pd
import numpy as numpy
import os
import xlrd
df = pd.DataFrame()

appended_data = []


for WorkingFile in os.listdir('C:/Users/Reube/Desktop/Excel File Assessment/Data Bank/'):


    x1_workbook = xlrd.open_workbook("Data Bank/"+ WorkingFile)

    first_sheet = x1_workbook.sheet_by_index(0)

    data = {
    'Expense Type 2': first_sheet.cell(5,0).value,
    'Expense Quantity 1': first_sheet.cell(4,1).value,
    'Expense Type 1': first_sheet.cell(4,0).value,
    'Expense Quantity 2': first_sheet.cell(5,1).value,
    'Expense Price 1': first_sheet.cell(4,2).value,
    'Expense Price 2': first_sheet.cell(5,2).value,
    'Revenue Type 1': first_sheet.cell(4,4).value,
    'Revenue Type 2': first_sheet.cell(5,4).value,
    'Revenue Quantity 1': first_sheet.cell(4,5).value,
    'Revenue Quantity 2': first_sheet.cell(5,5).value,
    'Revenue Price 1': first_sheet.cell(4,6).value,
    'Revenue Price 2': first_sheet.cell(5,6).value,
    'Net Revenue': first_sheet.cell(8,2).value,
    'Profit': first_sheet.cell(8,3).value,

    }
    df = df.append(data,ignore_index=True)
    print(df.describe())
df.to_csv('Treasurer_Budgets.csv')
