import traceback
import pytest
import os
import sys
import time
import datetime
o_path=os.path.abspath(os.path.join(os.getcwd(),'..'))
sys.path.append(o_path)
from U_LoginPage import LoginPage
from selenium import webdriver


# def test_Preservation(my_fixture):
#     personinfo=my_fixture[0]
def test_Preservation():
    try:
        options = webdriver.ChromeOptions()
        options.add_argument('--ignore-certificate-errors')
        driver = webdriver.Chrome(options=options)
        driver.implicitly_wait(30)

        url ="https://10.1.113.23:9443/"
        username ="TEST41"
        password ="admin001"

        # Policy_number='9005009993'# (女)9005010510


    ## ------登入---------
        test= LoginPage(driver, url, u"人壽保險核心業務系統")
        test.open()
        test.switch_frame('fraInterface')
        test.input_username(username)
        test.input_password(password)       
        test.click_submit()
        print('登入結束')

        test.mouse_To_NewIInput()
        time.sleep(1)
        test.switch_targetWindow("扫描件显示")
        test.NewInput_PolAppntDate_click()
        






    except Exception as e:
        print('有錯誤'+str(e))
        traceback.print_exc()
        driver.quit()

    finally:
        print('測試結束')
        driver.quit()



if __name__ == "__main__":
    pytest.main()