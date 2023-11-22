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
driver.get("https://passport.shop.jd.com/")

driver.find_element(By.XPATH, "/html/body/div/div/div/div[1]/div[2]").click()
time.sleep(1)
driver.find_element(By.XPATH, "/html/body/div/div/div/div[1]/div[2]").click()

driver.switch_to.frame('loginFrame')
time.sleep(2)
driver.find_element(By.ID, 'loginname').send_keys('壹时钟表海外专营店')
driver.find_element(By.XPATH, '//*[@id="nloginpwd"]').send_keys('c01010142')

driver.find_element(By.ID, 'paipaiLoginSubmit').click()



check = input('Enter "Y" after login')

if check == "Y":
    driver.get('https://fin.shop.jd.com/taurus/billManageIndex#/daybill')
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH,'/html/body/div[2]/div/div[2]/button[2]'))
        )
        driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/button[2]').click()

    finally:
        print('success!')

    driver.find_element(By.XPATH,'//*[@id="tab-statement"]').click()
    refList = []
    # for order in range(1,11):
    #     ref = driver.find_element(By.XPATH,'//*[@id="app"]/div[2]/div[3]/div[4]/div[3]/table/tbody/tr[{}]'.format(order))
    #     refList.append(ref.text)
    # print(refList)
    # ref = driver.find_element(By.XPATH,'//*[@id="app"]/div[2]/div[3]/div[4]/div[3]/table/tbody/tr[{}]'.format(1))
    while True:
        x = input('quit to quit')
        if x.lower() == 'quit':
            driver.quit()
            exit()

'//*[@id="app"]/div[2]/div[3]/div[4]/div[3]/table/tbody/tr[1]'
'//*[@id="app"]/div[2]/div[3]/div[4]/div[3]/table/tbody/tr[2]'