import os
import sys
import traceback
import pytest
import time

o_path=os.path.abspath(os.path.join(os.getcwd(),'..'))
sys.path.append(o_path)

from U_LoginPage import LoginPage
from selenium import webdriver
# from ArtiConfirm import test_artiConfirm

def test_autoConfirm(my_fixture):
    try:
        personinfo=my_fixture[0]
        options = webdriver.ChromeOptions()
        options.add_argument('--ignore-certificate-errors')
        driver = webdriver.Chrome(options=options)
        driver.implicitly_wait(30)
        url ="https://10.1.113.23:9443/"
        username ="TEST41"
        password ="admin001"

        test= LoginPage(driver, url, u"人壽保險核心業務系統")
        test.open()
        test.switch_frame('fraInterface')
        test.input_username(username)
        test.input_password(password)       
        test.click_submit()
        print('登入結束')

        # ------查詢保單號碼---------
        test.mouse_To_NewQuery()
        test.input_QueryPage_Date(personinfo)
        test.input_QueryPage_Name(personinfo)
        test.input_QueryPage_ID(personinfo)
        test.click_QueryPage_Button()
        test.check_QueryPage_seleOne()
        Policy_number=test.getNumbertext()

        test.mouse_To_Autoconfirm()
        test.input_Policy_number(Policy_number)
        test.click_Auto_Confirm_Search()
        test.click_Auto_Confirm_select()
        test.click_Auto_Confirm_Submit()

        # test_artiConfirm(Policy_number)

        time.sleep(5)
        driver.quit()

    except Exception as e:
        print('有錯誤'+str(e))
        traceback.print_exc()
        # driver.close()
        # driver.quit()
    finally:
        print('測試結束')

if __name__ == "__main__":
    pytest.main()