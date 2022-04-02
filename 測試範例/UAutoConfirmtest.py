
from U_LoginPage import LoginPage
from selenium import webdriver
import time

options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors')
driver = webdriver.Chrome(chrome_options=options)
driver.implicitly_wait(30)
url ="https://10.1.113.23:9443/"
username ="TEST41"
password ="admin001"
Date="2021-10-27"
Name="吳O鷹"
Policy_number='9000140185'

try:
    test1=LoginPage(driver, url, u"人壽保險核心業務系統")
    test1.open()
    test1.switch_frame('fraInterface')
    test1.input_username(username)
    test1.input_password(password)       
    test1.click_submit()
    print('登入結束')
    test1.mouse_To_Autoconfirm()
    test1.input_Policy_number(Policy_number)
    test1.click_Auto_Confirm_Search()
    test1.click_Auto_Confirm_select()

    test1.click_Auto_Confirm_Submit()

    time.sleep(5)

    driver.close()
    driver.quit()

except Exception as e:
    print('有錯誤'+str(e))
    driver.close()
    driver.quit()
finally:
    print('測試結束')