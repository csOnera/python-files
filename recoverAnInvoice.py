import mysql.connector
import os
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


while True:
    invoice = input('還原那一張invoice? or "quit" to quit')

    if invoice != "quit":
        os.system('cls')
        print(invoice)

        cursor.execute("""
            select `型號`, `數量` from refInvoiceNo
            where `發票` like '%{}%';    
        """.format(invoice))

        resultList = cursor.fetchall()
        # print(resultList)
        cursor.execute("""
            select `型號`, `數量`, `去處` from exportRecord
            where `發票` like '%{}%';
        """.format(invoice))

        resultList += cursor.fetchall()

        oneRefList = []
        keysList = []

        # item is each record
        for item in resultList:
            # will probably mark down all
            if item[0] not in keysList:
                oneRefList.append({item[0]:item[1], "去處": "", "總出數": 0})
                keysList.append(item[0])
            else:
                ind = keysList.index(item[0])
                if len(item) == 3:
                    oneRefList[ind]["去處"] += item[2] + ' -' + str(item[1]) +'\n'
                    oneRefList[ind][item[0]] += item[1]
                    oneRefList[ind]["總出數"] += item[1]
                else:
                    oneRefList[ind][item[0]] += item[1]


        SUM = 0
        LEFT = 0
        for i in oneRefList:
            keys = list(i.keys())
            SUM += i[keys[0]]
            LEFT += i[keys[0]] - int(i[keys[2]])
            print(keys[0]+ "  "*5 + str(i[keys[0]]) + " (" + str(i[keys[0]] - int(i[keys[2]])) + ")")
            displayList = i[keys[1]].split("\n")
            for j in range(len(displayList) - 1):
                print('`-----------------------------> ', end="")
                print(displayList[j])

        print(str(LEFT) + "/" + str(SUM))

    else:
        quit()
