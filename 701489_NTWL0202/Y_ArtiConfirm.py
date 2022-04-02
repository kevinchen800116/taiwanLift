import os
import sys
import traceback
import pytest
import time

o_path=os.path.abspath(os.path.join(os.getcwd(),'..'))
sys.path.append(o_path)

from U_LoginPage import LoginPage
from selenium import webdriver

def test_artiConfirm(my_fixture):
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
        # 移動到新單查詢
        test.mouse_To_NewQuery()
        test.input_QueryPage_Date(personinfo)
        test.input_QueryPage_Name(personinfo)
        test.input_QueryPage_ID(personinfo)
        test.click_QueryPage_Button()
        test.check_QueryPage_seleOne()
        Policy_number=test.getNumbertext()
        print('保單號碼：'+Policy_number)

        # 移動到人工核保
        test.mouse_To_Articalconfirm()

        # 輸入公共池內的保單號碼
        test.input_ArtiConfirmPage_first_PolicyNumber(Policy_number)

        # 點擊公共池內的查詢
        test.click_ArtiConfirmPage_first_QBtn()

        # 點擊公共池內的保單
        test.click_ArtiConfirmPage_PublicPool()

        # 輸入個人池內的保單號碼
        # test.input_ArtiConfirmPage_second_PoliceNumber(Policy_number)

        # 點擊個人池內的查詢
        # test.click_ArtiConfirmPage_second_Qbtn()

        # 點擊個人池內的保單
        # test.click_ArtiConfirmPage_second_PrivatePool()

        # 選擇五個照會
        # test.click_ArtiConfirmPage_check1to5()

        # 全選照會
        test.click_ArtiConfirmPage_checkAll()

        # 解除按鈕
        test.click_ArtiConfirmPage_AutoUWButton()

        # 選擇核保結論方法
        test.sele_ArtiConfirmPage_auwState()

        #選擇核保流向方法
        test.sele_ArtiConfirmPage_uwUpReport()

        # 輸入意見
        test.input_ArtiConfirmPage_UWIdea()

        # 點擊確定按鈕
        test.click_ArtiConfirmPage_button1()

        time.sleep(1)

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