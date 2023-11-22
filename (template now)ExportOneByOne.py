import mysql.connector

connection = mysql.connector.connect(
    host='localhost',
    port='3306',
    user='root',
    password='jdysz',
    database='trial_database'
)

cursor = connection.cursor()

ref = input('型號')
# check if there is this ref
cursor.execute("select `型號` from `refInvoiceNo` where `型號` = '{}';".format(ref))
check = cursor.fetchall()

if check == []:
    print("庫存沒有該型號")
    exit()
else:
    # check if enough stock
    takingNum = int(input('出多少(數量)'))
    to = input('去處')
    price = int(input('售價(整數)'))
    # ref = 'T120.417.11.051.00'
    # takingNum = 12
    # to = '22/5/2023 POP'

    cursor.execute("set @ref := '{}';".format(ref))
    cursor.execute("set @takingNum := '{}';".format(takingNum))


    # if not enough stock, show total number in stock

    cursor.execute("select `數量` from `refInvoiceNo` where `型號` = '{}';".format(ref))
    noList = cursor.fetchall()

    if len(noList) > 1:
        totalNum = 0
        for i in noList:
            totalNum += i[0]
    else:
        totalNum = noList[0][0]

    # print('totalNum: ',totalNum)
    if takingNum > totalNum:
        print('not enough stock, {} left.'.format(totalNum))
    else:
        cursor.execute("select * from `refInvoiceNo` where `型號` = '{}';".format(ref))
        refInfo = cursor.fetchall()
        for i in refInfo:
            print(i)
        while takingNum > 0:
            cursor.execute("set @num := (select `數量` from `refInvoiceNo` where `數量` <> 0 and `型號` = '{}' order by `id` asc limit 1);".format(ref))
            cursor.execute("set @id := (select `id` from `refInvoiceNo` where `數量` <> 0 and `型號` = '{}' order by `id` asc limit 1);".format(ref))
            cursor.execute("select @num;")
            num = cursor.fetchall()[0][0]

            if takingNum < num:
                #  minus number in refInvoiceNo table
                cursor.execute("update `refInvoiceNo` set `數量` = '{}' where `id` = @id;".format(num - takingNum))
                #  record in exportRecord table
                cursor.execute("insert into `exportRecord` (`日期`,`型號`,`發票`,`數量`, `去處`, `price`) values (curdate(), @ref, (select `發票` from `refInvoiceNo` where `id` = @id), '{}', '{}', '{}');".format(takingNum,to,price))
                takingNum = 0
            else:
                cursor.execute("insert into `exportRecord` (`日期`,`型號`,`發票`,`數量`, `去處`, `price`) values (curdate(), @ref, (select `發票` from `refInvoiceNo` where `id` = @id), (select `數量` from `refInvoiceNo` where `id` = @id), '{}','{}');".format(to,price))
                cursor.execute("update `refInvoiceNo` set `數量` = 0 where `id` = @id;")
                takingNum -= num
        print('succeeded!!!!!')
			
     



cursor.close()
connection.commit()
connection.close()

