import mysql.connector
import os


connection = mysql.connector.connect(
    host='localhost',
    port='3306',
    user='root',
    password='jdysz',
    database='trial_database'
)

cursor = connection.cursor()

while True:
    q = input('choose 全 records of 庫存(stock), 出貨紀錄(export) or 退貨紀錄(back)')
    os.system('cls')
    match q:
        case "stock":
            cursor.execute("""
                select * from `refInvoiceNo`;
            """)
        case "export":
            cursor.execute("""
                select * from `exportRecord`;
            """)
        case "back":
            cursor.execute("""
                select * from `退貨紀錄`;
            """)

    list = cursor.fetchall()

    for i in list:
        print(i)

# cursor.close()
# connection.close()