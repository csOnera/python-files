import mysql.connector
import openpyxl
import win32com.client
import re


import time

import os
from dotenv import load_dotenv

load_dotenv()

print("Loading...")

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

# get the list from dewuCAP for verification of changing changed name
cursor.execute("""
    select * from `dewuCAP`;
               """)
dewuCAPList = cursor.fetchall()
checkingList = []
for i in dewuCAPList:
    checkingList.append(i[1])
# print(checkingList)

# find dd/mm/yyyy 得物 in exportRecord

cursor.execute("""
    select `去處`, sum(`數量`) from `exportRecord`
    where `去處` like '%得物%'
    group by `去處`;
""")

result = cursor.fetchall()
print("current status 現時情況")
print("編號    出貨名    淨出貨數量    對應單號")
index = 0
haveCAPindexs = {}
for i in result:
    for j in range(len(checkingList)):
        if i[0] == checkingList[j]:
            print(index, i[0], i[1], dewuCAPList[j][0])
            haveCAPindexs[index] = j
            break
        elif i[0] != checkingList[j] and j == len(checkingList) - 1:
            print(index, i[0], i[1])
    index += 1

haveCAPindexsKeys = list(haveCAPindexs.keys())
changingChoice = int(input("請輸入想加對應單號的編號: "))

# two ways for adding or editing respectively

# check if there is CAP already
for i in range(len(haveCAPindexsKeys)):
    if changingChoice == haveCAPindexsKeys[i]:
    # first, editting the CAP
        verify = input("如確定要更改對應單號, 輸入'Y': ")
        if verify == 'Y':
            print(result[changingChoice])
            changeTo = input("輸入新對應單號: ")
            cursor.execute(f"""update `dewuCAP` set `CAP號` = '{changeTo}' where (`CAP號` = '{dewuCAPList[haveCAPindexs[changingChoice]][0]}');""")
    # second, adding a new insert in dewuCAP
    elif changingChoice != haveCAPindexsKeys[i] and i == len(haveCAPindexsKeys) - 1:
        print(result[changingChoice])
        changeTo = input("輸入新對應單號: ")
        cursor.execute(f"""insert into `dewuCAP` (`CAP號`, `exportRef`) values ('{changeTo}','{result[changingChoice][0]}');""")






cursor.close()
connection.commit()
connection.close()

print('成功修改!')

time.sleep(5)