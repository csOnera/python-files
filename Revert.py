# this program aims to revert the operation in inputExport by any means
# 

import openpyxl
import mysql.connector
import win32com.client
from exportExport import exportExport, runVBA


connection = mysql.connector.connect(
    host='localhost',
    port='3306',
    user='root',
    password='jdysz',
    database='trial_database'
)

cursor = connection.cursor()

# two parts
# first to add back the number in `refInvoiceNo`
# get the invoice and ref from `exportRecords`

needExport = input('is export record done? "S" to skip exporting again')
if needExport != "S":
    ref = input('請輸入要查找型號或輸入"quit"退出查找')

    if ref == "quit":
        quit()
    else:
        idList = []
        cursor.execute("""
            select * from `exportRecord`
            where `型號` = '{}';
        """.format(ref))
        list = cursor.fetchall()
        if list == []:
            print('庫存沒有該記錄\n')
        for i in list:
            print(i)
            idList.append(i[0])
    undoId = 0
    while undoId not in idList:
        undoId = int(input("id (integer)"))

    exportExport("id",skipAskingId=undoId)
    
xl = win32com.client.Dispatch("Excel.Application")
xl.Visible = True

wb_open = True
wb = xl.Workbooks.Open(r"C:\Users\onera\OneDrive - ONE ERA (HK) LIMITED\oneraShare\DATABASE_TRIAL\python files\exportExportRecords.xlsm")


checked = input('CHECK items to be reverted in the exportExport.xlsm "checked"')
# while True:
if checked == "checked" :
    wb.Close(False)
    excel = openpyxl.load_workbook(r"C:\Users\onera\OneDrive - ONE ERA (HK) LIMITED\oneraShare\DATABASE_TRIAL\python files\exportExportRecords.xlsm")
    sheet = excel.active
    # loop for >1 records to be reverted in .xlsm
    for i in range(2, sheet.max_row + 1):
        exportId = sheet["a" + str(i)].value
        # ref = sheet["c" + str(i)].value
        # invoice = sheet["d" + str(i)].value
        # to = sheet["f" + str(i)].value
        numAddBack = sheet["e" + str(i)].value
        refId = sheet["h" + str(i)].value


        # take the no. in stock first (big bug happened: more than one record of same ref same invoice)
        # should take ref_id
        cursor.execute("""
            select `數量` from `refInvoiceNo`
            where `id` = '{}';
        """.format(refId))
        NumInStock = cursor.fetchall()[0][0]
        cursor.execute("""
            update `refInvoiceNo`
            set `數量` = '{}'
            where `id` = '{}';
        """.format(NumInStock + numAddBack, refId))

        # second to delete the records in `exportRecords`???
        cursor.execute("""
            delete from `exportRecord`
            where `id` = '{}'
        """.format(exportId))


print("done!")
xl.Application.Quit()

import RefreshStock
RefreshStock
import RefreshRecord
RefreshRecord

cursor.close()
connection.commit()
connection.close()