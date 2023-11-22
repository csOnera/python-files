import mysql.connector

connection = mysql.connector.connect(
    host='localhost',
    port='3306',
    user='root',
    password='jdysz',
    database='trial_database'
)
cursor = connection.cursor()

# def fillAllRefId():
    # cursor.execute("select * from `exportRecord`;")
    # allList = cursor.fetchall()

    # for i in allList:
    #     id = i[0]
    #     ref = i[2]
    #     invoice = i[3]
    #     cursor.execute("""
    #             select `id` from `refInvoiceNo`
    #             where `發票` = '{}' and `型號` = '{}'
    #         """.format(invoice, ref))
    #     ref_id = cursor.fetchall()[0][0]
    #     cursor.execute("""
    #         update `exportRecord`
    #         set `ref_id` = '{}'
    #         where `id` = '{}';
    #     """.format(ref_id, id))
    

cursor.close()
connection.commit()
connection.close()

import openpyxl
import pathlib

# folder = pathlib.Path(r"C:\Users\onera\OneDrive - ONE ERA (HK) LIMITED\oneraShare\得物對賬\23 3.01-5.28")
# flist = list(folder.iterdir())
# for i in range(4, len(flist)):

# print(flist[3])
# for j in range(4, len(flist)):
#     excel = openpyxl.load_workbook(flist[j])
#     sheet = excel["鉴定通过订单"]

#     for i in range(4, sheet.max_row + 1):
#         sheet["b" + str(i)].value = "'" + sheet["a" + str(i)].value

#     excel.save(flist[j])


    # print(sheet["b4"].value)

# from AppOpener import open
# open("京麦")

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
# from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.common.by import By

from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# service_obj = Service("chromedriver.exe")
# driver = webdriver.Chrome(service=service_obj)
# driver.get("http://amc.wtdex.com/amc/login;jsessionid=D7DAF284BAA77F5858E9945D3FC2C1CA")


# from PIL import Image
# import pytesseract

# image = driver.find_element(By.XPATH,'//*[@id="VCode"]')

# src = image.get_attribute('src')
# import requests 
# from io import BytesIO
# response = requests.get(src)
# img = Image.open(BytesIO(response.content))

# text = pytesseract.image_to_string(Image.open("D:\Downloads\images.jfif"))
# print(text)
# driver.get("https://www.google.com/search?rlz=1C1ONGR_zh-HKHK950HK951&q=table&si=AMnBZoEEPkISis2GwXGfNE5GFpR2pnfLfcCPgv2IkQ17iDSEyc-7oXoq1tcGjv227XHS6FTLkA6IU581Zg9xjTf_L7oB9n4Knw%3D%3D&expnd=1&sa=X&ved=2ahUKEwiev4qotLP_AhU9iO4BHY0HCW8Q2v4IegUIFRDDAQ&biw=1920&bih=350&dpr=1")

# driver.find_element(By.XPATH,'//*[@id="tsuid_43"]/span/div/div/div[3]/div/div[2]/div[2]/div/ol')

# while True:
#     pass

# dicti = {'m123': 2, 'm5641': 3}

# print(dicti[0])

