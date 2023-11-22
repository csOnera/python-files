from selenium import webdriver
from selenium.webdriver.chrome.service import Service
# from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.common.by import By

from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# import selenium.webdriver.support.ui as ui

# from webdriver_manager.chrome import ChromeDriverManager
import time

# ADD ID, MONEY
service = Service()
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)
driver.get("https://mms.pinduoduo.com/")


driver.find_element(By.XPATH, '//*[@id="root"]/div[1]/div/div/main/div/section[2]/div/div/div/div[1]/div/div[2]').click()

time.sleep(2)
driver.find_element(By.ID, 'usernameId').send_keys('15011232105')
driver.find_element(By.ID,'passwordId').send_keys('Op123456')

# driver.find_element(By.XPATH, '//*[@id="root"]/div[1]/div/div/main/div/section[2]/div/div/div/div[2]/section/div/div[2]/button').click()


check = input('Enter "Y" after login')

if check == "Y":
    driver.get('https://mms.pinduoduo.com/finance/payment-bills/index?q=1')
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH,'//*[@id="mf-mms-finance-web-container"]/div/div/div[3]/div/div[2]/div/div[1]/div[1]/div/div/div/div/div/div/div[1]/input'))
        )
        driver.find_element(By.XPATH,'//*[@id="mf-mms-finance-web-container"]/div/div/div[3]/div/div[2]/div/div[1]/div[1]/div/div/div/div/div/div/div[1]/input').send_keys('2023-09-01 00:00:00 ~ 2023-09-30 23:59:59')
    finally:
        print('success!')

    while True:
        x = input('quit to quit')
        if x.lower() == 'quit':
            driver.quit()
            exit()