
from U_LoginPage import LoginPage
from selenium import webdriver


options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors')
driver = webdriver.Chrome(chrome_options=options)
driver.implicitly_wait(30)
url ="https://10.1.113.23:9443/"
username ="TEST41"
password ="admin001"
Date="2021-10-27"
Name="吳O鷹"

try:
    test= LoginPage(driver, url, u"人壽保險核心業務系統")
    test.open()
    test.switch_frame('fraInterface')
    test.input_username(username)
    test.input_password(password)       
    test.click_submit()
    print('登入結束')
    test.mouse_To_NewQuery()
    test.input_queryDate(Date)
    test.input_queryName(Name)
    test.click_queryButton()
    test.check_seleOne()
    Policy_number=test.getNumbertext()
    print(Policy_number)
    driver.close()
    driver.quit()

except Exception as e:
    print('有錯誤'+str(e))
    driver.close()
    driver.quit()
finally:
    print('測試結束')