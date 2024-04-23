import mysql.connector
import openpyxl
import win32com.client
import re


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

def matchInsideSame(ref, exportList):
    for row in exportList:
        if row in ref:
            return exportList.index(row)
    return False

while True:
    
    cap = input("請輸入要查看的CAP/AP: ")
    cursor.execute(f"""
        select `型號`, count(`引號`) as 數量, `對賬單號` from `得賬`
        where `CAP號` = '{cap}'
        group by `型號`, `對賬單號`
        order by `型號`;
    """)
    dewu = cursor.fetchall()

    cursor.execute(f"""
        select `型號`, sum(數量) as 寄數量 from `exportRecord`
        where 去處 = (
            select exportRef from dewuCAP 
            where CAP號 = '{cap}'
        )
        group by `型號`
        order by `型號`;
    """)
    export = cursor.fetchall()

    # print(len(dewu), len(export))
    if len(dewu) == 0 and len(export) == 0:
        print(f"沒找到相關紀錄: {cap}")

    

    # main stuff
    elif len(dewu) != 0 and len(export) != 0:
        
        wholeDict = []
        refList = []
        nomatch = []
        sumNo = 0
        soldSum = 0

        for row in export:
            wholeDict.append({row[0]: int(row[1]), '已售數量': 0})
            sumNo += int(row[1])
            refList.append(row[0])

        for row in dewu:
            matchingIndex = matchInsideSame(row[0], refList)

            if matchingIndex:
                # checkout if there is record first
                if 'records' in wholeDict[matchingIndex].keys():
                    wholeDict[matchingIndex]['records'].append(row[2] + " -" + str(row[1]))
                    wholeDict[matchingIndex]['已售數量'] += int(row[1])
                    soldSum += int(row[1])
                else:
                    wholeDict[matchingIndex]['records'] = [row[2] + " -" + str(row[1])]
                    wholeDict[matchingIndex]['已售數量'] = int(row[1])
                    soldSum += int(row[1])
            else:
                nomatch.append(list(row))
        for item in wholeDict:
            print(list(item.keys())[0] + "     " + str(item[list(item.keys())[0]]) + "("+ str(item[list(item.keys())[0]] - item['已售數量']) +")")
            if 'records' in item.keys():
                for record in item['records']:
                    print("`-----------> " + record)
        for item in nomatch:
            print(item)
        print(str(soldSum) +"/" + str(sumNo))
        
        


    # end
    elif len(export) != 0:
        print(f"只有出貨的紀錄: {cap}")
        for row in export:
            print(row)
    else:
        print(f"只有對賬的紀錄: {cap}")
        for row in dewu:
            print(row)


    # if len(dewu) == 0:
    #     print(f"沒找到相關紀錄: {cap}")
    # else:
    #     for row in dewu:
    #         print(row)

