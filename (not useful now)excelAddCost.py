# This file tries to insert data(cost) from excel to database

import mysql.connector
import openpyxl

# path = "C:\\Users\\onera\\OneDrive - ONE ERA (HK) LIMITED\\oneraShare\\DATABASE_TRIAL\\refInvoiceNo.xlsx"


excel = openpyxl.load_workbook("refInvoiceNo.xlsx")

sheet = excel.active



connection = mysql.connector.connect(
    host='localhost',
    port='3306',
    user='root',
    password='jdysz',
    database='trial_database'
)

cursor = connection.cursor()

for i in range(50,209):
    if sheet['f'+str(i)].value != None:
        cost = sheet['f'+ str(i)].value
        id = sheet['a' + str(i)].value

        cursor.execute("update `refInvoiceNo` set `cost` = '{}' where `id` = '{}';".format(cost, id))
    else:
        pass
# print(sheet['g47'].value)



cursor.close()
connection.commit()
connection.close()