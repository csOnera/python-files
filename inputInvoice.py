# pre-requisite: data inputted in InputInvoice.xlsx
# this program will add the data of invoice into RefInvoiceNo table in sql
# Since it wont have too much add/drops within tables, 
# confirmation question is not set
# (wont add records in `invoicDate` table currently)

import mysql.connector
import openpyxl
import win32com.client

xl = win32com.client.Dispatch("Excel.Application")
xl.Visible = True
wb_open = True
wb = xl.Workbooks.Open(r"C:\Users\onera\OneDrive - ONE ERA (HK) LIMITED\oneraShare\DATABASE_TRIAL\python files\InputInvoice.xlsx")



confirm = input(' EXCEL sheet entered? "Y"')

if confirm == "Y":
    wb.Close(False)
    # below line new added
    xl.Application.Quit()

    excel = openpyxl.load_workbook(r"C:\Users\onera\OneDrive - ONE ERA (HK) LIMITED\oneraShare\DATABASE_TRIAL\python files\InputInvoice.xlsx")
    sheet = excel.active


    connection = mysql.connector.connect(
        host='localhost',
        port='3306',
        user='root',
        password='jdysz',
        database='trial_database'
    )

    cursor = connection.cursor()


    for i in range(2, sheet.max_row + 1):
        ref = sheet['a' + str(i)].value
        if ref == None:
            print(i, sheet.max_row)
            break
        invoice = sheet['b' + str(i)].value   
        num = int(sheet['c' + str(i)].value)
        現退 = sheet['d' + str(i)].value
        cost = float(sheet['e' + str(i)].value)
        csorone = sheet['f' + str(i)].value


        cursor.execute("""
        insert into `refInvoiceNo` (`型號`, `發票`, `數量`, `現貨/退貨/盒`, `cost`, `cs/onera`, `入貨日期`) 
        values ('{}', '{}', '{}', '{}', '{}', '{}', curdate());
                       
        """.format(
            ref, invoice, num, 現退, cost, csorone
        ))

    print("succeeded!!!!")
        

    excel.save(r"C:\Users\onera\OneDrive - ONE ERA (HK) LIMITED\oneraShare\DATABASE_TRIAL\python files\InputInvoice.xlsx")
    
    cursor.close()
    connection.commit()
    connection.close()

import RefreshStock
RefreshStock

while True:
    pass
