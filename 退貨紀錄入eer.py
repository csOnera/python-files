import mysql.connector
import openpyxl
import win32com.client
import os


connection = mysql.connector.connect(
    host='localhost',
    port='3306',
    user='root',
    password='jdysz',
    database='trial_database'
)

cursor = connection.cursor()


eer = openpyxl.load_workbook(r"C:\Users\onera\OneDrive - ONE ERA (HK) LIMITED\oneraShare\DATABASE_TRIAL\python files\exportExportRecords.xlsm", read_only=False, keep_vba=True)
eerws = eer.active
    

# print col name
cursor.execute("""
    show columns from `exportRecord`;
""")
records = cursor.fetchall()
records.append(["cost"])
records.append(["現貨/退貨/盒"])
colNum = len(records) 
for i in records:
    eerws.cell(row = 1, column = records.index(i) + 1).value = i[0]

# export the last 退貨紀錄
cursor.execute("""
    select `入貨單號`, `型號`, `backNum`, `出貨單號` from `退貨紀錄`
    order by `id` desc limit {};
""".format(1))
result = cursor.fetchall()

for row in range(2, len(result) + 2):
    eerws["c" + str(row)].value = result[row-2][1]
    eerws["d" + str(row)].value = result[row-2][0]
    eerws["e" + str(row)].value = result[row-2][2]
    eerws["f" + str(row)].value = result[row-2][3] + "退貨 +"

eer.save(r"C:\Users\onera\OneDrive - ONE ERA (HK) LIMITED\oneraShare\DATABASE_TRIAL\python files\exportExportRecords.xlsm")
from exportExport import runVBA

runVBA()

import RefreshStock
RefreshStock
import RefreshRecord
RefreshRecord