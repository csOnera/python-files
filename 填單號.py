from selenium import webdriver
from selenium.webdriver.chrome.service import Service
# from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.common.by import By

# from selenium.webdriver.support.wait import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# import selenium.webdriver.support.ui as ui

# from webdriver_manager.chrome import ChromeDriverManager
import time

# ADD ID, MONEY
service = Service()
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)
driver.get("https://passport.shop.jd.com/")


driver.find_element(By.XPATH, "/html/body/div/div/div/div[1]/div[2]").click()
time.sleep(1)
driver.find_element(By.XPATH, "/html/body/div/div/div/div[1]/div[2]").click()

driver.switch_to.frame('loginFrame')
time.sleep(2)
driver.find_element(By.ID, 'loginname').send_keys('壹时钟表-发货')
driver.find_element(By.XPATH, '//*[@id="nloginpwd"]').send_keys('OP123456')

driver.find_element(By.ID, 'paipaiLoginSubmit').click()


check = input('Enter "Y" after login')

if check == "Y":
    driver.get("https://porder.shop.jd.com/order/orderlist/waitInnerShip?t=1687839370895")

table = driver.find_element(By.XPATH,'//*[@id="order-shop-content"]/div/div/div/div[7]/div[2]')
'//*[@id="order-shop-content"]/div/div/div/div[7]/div[2]'
count = len(table.find_elements(By.XPATH,'.//div/div/div[1]/table/thead/tr/th/div[1]/span[1]'))

print(count)
infoList = []

import re

for i in range(1, count + 1):
    note = table.find_element(By.XPATH, './/div[{}]/div/div[2]/table/tbody/tr/td[6]'.format(i)).text
    print(note)
    if note == '':
        ref = table.find_element(By.XPATH,'.//div[{}]/div/div[2]/table/tbody/tr/td[1]/div/div/p[1]/a'.format(i)).text
        ref = re.search("[a-zA-Z.\d\-]+\Z", ref).group()
        訂單號 = table.find_element(By.XPATH, './/div[{}]/div/div[1]/table/thead/tr/th/div[1]/span[1]/a'.format(i)).text
        price = float(table.find_element(By.XPATH, './/div[{}]/div/div[2]/table/tbody/tr/td[3]/p[2]'.format(i)).text[1:])
        infoList.append([ref, 訂單號, price])
        '//*[@id="order-shop-content"]/div/div/div/div[7]/div[2]/ div[2]/div/div[1]/table/thead/tr/th/div[1]/label/span'
        '//*[@id="order-shop-content"]/div/div/div/div[7]/div[2]/ div[2]/div/div[1]/table/thead/tr/th/div[1]/span[1]'
        '//*[@id="order-shop-content"]/div/div/div/div[7]/div[2]/div[3]/div/div[1]/table/thead/tr/th/div[1]/span[1]'

print(infoList)

driver.close()


# 訂單號:  /div[x]/div/div[1]/table/thead/tr/th/div[1]/span[1]/a
# ref:   /div[1]/div/div[2]/table/tbody/tr/td[1]/div/div/p[1]/a
# note:  /div[1]/div/div[2]/table/tbody/tr/td[6]/div/div[1]/a
# //*[@id="order-shop-content"]/div/div/div/div[6]/div[2]/div[1]/div/div[2]/table/tbody/tr/td[6]

import openpyxl
from datetime import datetime
import win32com.client
xl = win32com.client.Dispatch("Excel.Application")

wb = xl.Workbooks.Open(r"C:\Users\onera\OneDrive - ONE ERA (HK) LIMITED\oneraShare\DATABASE_TRIAL\python files\exportExportRecords.xlsm")
POP = openpyxl.load_workbook(r"C:\Users\onera\OneDrive - ONE ERA (HK) LIMITED\oneraShare\出貨OR退貨紀錄\POP出貨記錄LATEST VERSION-DESKTOP-833R29B.xlsx")
popws = POP.active

for i in range(1500, popws.max_row + 1):
    # find the stopper
    if popws["q" + str(i)].value != None:
        j = i
        while True:
            if popws['e' + str(j)].value != None:
                print(j)
                break
            j -= 1
        # add date first
        popws['a' + str(j+2)].value = datetime.today().strftime('%d/%m/%Y')
        
        # popws.insert_rows(j + 1, len(infoList))
        # add info
        for k in range(len(infoList)):
            popws['b' + str(j+2+k)].value = infoList[k][1]
            popws['e' + str(j+2+k)].value = infoList[k][0]
            popws['g' + str(j+2+k)].value = infoList[k][2]
            
        break

POP.save(r"C:\Users\onera\OneDrive - ONE ERA (HK) LIMITED\oneraShare\出貨OR退貨紀錄\POP出貨記錄LATEST VERSION-DESKTOP-833R29B.xlsx")
pop = xl.Workbooks.Open(r"C:\Users\onera\OneDrive - ONE ERA (HK) LIMITED\oneraShare\出貨OR退貨紀錄\POP出貨記錄LATEST VERSION-DESKTOP-833R29B.xlsx")
xl.Application.Run('exportExportRecords.xlsm!Module2.insertRowInPOP')
print('succeeded!!')

pop.Close(True)

xl.Application.Quit()


while True:
    ans = input('quit if done')
    if ans == 'quit':
        quit()