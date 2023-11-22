import mysql.connector

connection = mysql.connector.connect(
    host='localhost',
    port='3306',
    user='root',
    password='jdysz',
    database='trial_database'
)

cursor = connection.cursor()

def 全庫存():
    cursor.execute("""
        select * from `refInvoiceNo`;
    """)
    list = cursor.fetchall()
    for i in list:
        print(i)


def 冇零全庫存():
    cursor.execute("""
        select * from `refInvoiceNo`
        where `數量` <> 0;
    """)
    list = cursor.fetchall()
    for i in list:
        print(i)


def 全出貨紀錄():
    cursor.execute("""
        select * from `exportRecord`;
    """)
    list = cursor.fetchall()
    for i in list:
        print(i)

def 庫存價值():
    cursor.execute("""
        select sum(`cost`) from `refInvoiceNo`;
    """)
    i = cursor.fetchall()[0][0]
    print(i)


def 查找庫存中一個型號(ref):
    cursor.execute("""
        select * from `refInvoiceNo`
        where `型號` = '{}'
        and `數量` <> 0;
    """.format(ref))
    list = cursor.fetchall()
    for i in list:
        print(i)

def 查找最近的多少紀錄(num):
    cursor.execute("""
        select `exportRecord`.`id`, `exportRecord`.`型號`, `exportRecord`.`數量`, `exportRecord`.`發票`, `exportRecord`.`去處`, 
        `price`, `cost` from `exportRecord`
        join `refInvoiceNo` 
        on `ref_id` = `refInvoiceNo`.`id`
        order by `exportRecord`.`id` desc limit {};
    """.format(num))
    list = cursor.fetchall()
    list.reverse()
    for i in list:
        print(i[6])


def 查找某一訂單的型號紀錄(單號):
    cursor.execute("""
        select `型號`,`發票`,`數量`, `去處`, price from `exportRecord`
        where `去處` like '%{}%';
    """.format(單號))
    list = cursor.fetchall()
    for i in list:
        print(i)

def 更新最新庫存():
    import RefreshStock
    RefreshStock

def 輸入新存貨():
    import inputInvoice
    inputInvoice

def 輸入出貨紀錄():
    import inputExport
    inputExport

def 取消紀錄並加回庫存():
    import Revert
    Revert

def 自動約倉excel():
    import AUTO入倉
    AUTO入倉


# Examples:
# 查找庫存中一個型號('h77765541')
# 查找最近的多少紀錄(4)
# 查找某一訂單的型號紀錄('C024')

# 取消紀錄並加回庫存()
# 全出貨紀錄()
# 冇零全庫存()
# 全庫存()
# 庫存價值()

# 輸入新存貨()
# 自動約倉excel()
輸入出貨紀錄()



cursor.close()
connection.commit()
connection.close()
更新最新庫存()