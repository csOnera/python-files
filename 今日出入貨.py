# version 2: fixed bug about tdy exporting yesterdays stock (which will minus the stock of yesterdays) => one more check in tmrs export and add back to todaysStock list

import mysql.connector
import os

def printing(date, determinator):
    connection = mysql.connector.connect(
        host='localhost',
        port='3306',
        user='root',
        password='jdysz',
        database='trial_database'
    )

    cursor = connection.cursor()

    todaysInvoice = []
    cursor.execute("""
        select * from `refInvoiceNo` 
        where `入貨日期` = {};
    """.format(date))  
    # get todays invoice
    result = cursor.fetchall()
    for item in result:
        if item[2] not in todaysInvoice:
            todaysInvoice.append(item[2])

    todaysStock = []

    # add exported ones in a list to be added back to todays stock

    exportList = []
    cursor.execute("""
        select id, 型號, 發票, 數量, 去處, price from `exportRecord`
        where `日期` = {};
    """.format(date))
    list = cursor.fetchall()

    cursor.execute("""
        select sum(`數量`) from `exportRecord`
        where `日期` = {};
    """.format(date))
    exportTotalNum = cursor.fetchall()[0][0]

    # add tmrexportRecords
    cursor.execute("""
        select id, 型號, 發票, 數量, 去處, price from `exportRecord`
        where `日期` > {};
    """.format(date))
    tmrExportRecords = cursor.fetchall()

    for item in tmrExportRecords:
        if item[2] in todaysInvoice:
            todaysStock.append(item)


    if list == []:
        print(f'{"今日" if determinator == 1 else "昨日"}未有出貨\n')
    else:
        for i in list:
            if i[2] not in exportList:
                exportList.append(i[2])
            if i[2] in todaysInvoice:
                todaysStock.append(i)
            print(i)

    print(f'{"今日" if determinator == 1 else "昨日"}出貨總數: ', end="")
    print(exportTotalNum) if exportTotalNum != None else print(0)

    print("\n========================================================\n")

    # check if lot sum = 0 as REMINDER
    for lot in exportList:
        cursor.execute(f"""
            select sum(`數量`) from `refInvoiceNo`
            where `發票` = '{lot}';   
        """)
        lotNum = cursor.fetchall()[0][0]
        if lotNum == 0:
            print("該訂單現已歸零: " + str(lot))
        else:
            print(str(lot) + "現還剩" + str(lotNum))

    print("\n========================================================\n")


    # 退貨
    cursor.execute("""
        select 型號, 入貨單號, backNum, 出貨單號, new_refId from `退貨紀錄`
        where `input_date` >= {};
    """.format(date))
    list = cursor.fetchall()

    cursor.execute("""
        select sum(`backNum`) from `退貨紀錄`
        where `input_date` >= {};
    """.format(date))
    exportTotalNum = cursor.fetchall()[0][0]

    if list == []:
        print(f'{"今日" if determinator == 1 else "昨日"}未有退貨入\n')
    else:
        for i in list:
            print(i)

    print(f'{"今日" if determinator == 1 else "昨日"}退貨總數: ', end="")
    print(exportTotalNum) if exportTotalNum != None else print(0)

    print("\n========================================================\n")
    # 入貨
    cursor.execute("""
        select id, 型號, 發票, 數量, `現貨/退貨/盒`, cost, `cs/onera` from `refInvoiceNo`
        where `入貨日期` = {};
    """.format(date))
    list = cursor.fetchall()

    # cursor.execute("""
    #     select sum(`數量`) from `refInvoiceNo`
    #     where `入貨日期` = {};
    # """.format(date))
    # exportTotalNum = cursor.fetchall()[0][0]
 
    count = 0
    if list == []:
        print(f'{"今日" if determinator == 1 else "昨日"}未有入貨\n')
    else:
        for i in list:
            count += i[3]
            print(i)
        for i in todaysStock:
            count += i[3]
            print(i)

    print(f'{"今日" if determinator == 1 else "昨日"}入貨總數 (不含退貨): ', end="")
    print(count) if count != None else print(0)




    cursor.close()
    connection.close()

determinator = 1

def question(d):
    q = input("enter to switch to yesterday's or today's records")
    if q == '':
        return d * -1



while True:
    os.system('cls')
    if determinator == 1:
        printing('curdate()',determinator)
        
    elif determinator == -1:
        printing('DATE_SUB(CURDATE(), INTERVAL 1 DAY)',determinator)

    determinator = question(determinator)