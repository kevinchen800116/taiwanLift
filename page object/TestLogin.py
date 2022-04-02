from LoginPage import LoginPage
from BasePage import BasePage
from selenium import webdriver
import time


options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors')
driver = webdriver.Chrome(chrome_options=options)
driver.implicitly_wait(30)
url ="https://10.1.113.23:9443/"
username ="TEST41"
password ="admin001"

try:
    #声明LoginPage类对象
    login_page = LoginPage(driver, url, u"人壽保險核心業務系統")
    #调用打开页面组件
    login_page.open()
    #切换到登录框Frame
    login_page.switch_frame('fraInterface')
    #调用用户名输入组件
    login_page.input_username(username)    
    #调用密码输入组件
    login_page.input_password(password)       
    #调用点击登录按钮组件
    login_page.click_submit()
    time.sleep(5)
    driver.close()
    driver.quit()
except:
    print('有錯誤')
    driver.close()
    driver.quit()
finally:
    print('測試結束')
