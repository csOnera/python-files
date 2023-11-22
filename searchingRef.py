import mysql.connector

connection = mysql.connector.connect(
    host='localhost',
    port='3306',
    user='root',
    password='jdysz',
    database='trial_database'
)

cursor = connection.cursor()

while True:
    ref = input('請輸入要查找型號或輸入"quit"退出查找')

    if ref == "quit":
        break
    else:
        cursor.execute("""
            select * from `refInvoiceNo`
            where `型號` like '{}%';
        """.format(ref))
        list = cursor.fetchall()
        if list == []:
            print('庫存沒有該記錄\n')
        for i in list:
            print(i)



cursor.close()
connection.close()

quit()