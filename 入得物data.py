
import mysql.connector
import openpyxl
import win32com.client


import os
from dotenv import load_dotenv

load_dotenv()

MYSQL_USER = os.getenv('MYSQL_USER')
MYSQL_PW = os.getenv('MYSQL_PW')

connection = mysql.connector.connect(
    host='localhost',
    port='3306',
    user= MYSQL_USER,
    password= MYSQL_PW,
    database='trial_database'
)

cursor = connection.cursor()

# here enter the loading sheet
wb_path = r"C:\Users\onera\OneDrive - ONE ERA (HK) LIMITED\oneraShare\得物對賬//" + input("please input the 交易成功 name\ne.g. xxx.xlsx : ")

excel = openpyxl.load_workbook(wb_path)
ws = excel.active

for i in range(2, ws.max_row):
    if ws["c" + str(i)].value != None and ws["w" + str(i)].value != None:
        # bug by 引號 lol
        引號 = ws["c" + str(i)].value[1:]
        型號 = ws["h" + str(i)].value
        CAP號 = ws["t" + str(i)].value
        售價HKD = float(ws["n" + str(i)].value)
        對賬單號 = ws["w" + str(i)].value

        # print("""insert into `得賬` (`引號`, `型號`, `CAP號`, `售價HKD`, `對賬單號`) values 
        #       ('{}', '{}', '{}', {});"""
        #     .format(
        #         引號, 型號, CAP號, 售價HKD
        #     ))

        cursor.execute("""insert into `得賬` (`引號`, `型號`, `CAP號`, `售價HKD`, `對賬單號`) 
                       VALUES ('{}', '{}', '{}', {}, '{}');
                       """.format(
                           引號, 型號, CAP號, 售價HKD, 對賬單號
                       ))
print("data added")

excel.save(wb_path)

cursor.close()
connection.commit()
connection.close()