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
        username ="TEST26"
        password ="admin002"

        # Policy_number='9005010326'# (女)
        # Policy_number='9005010105'# (男)
        # Policy_number='9005010325'# (女)
        # Policy_number='9005010327'# (女)

        # Policy_number='9005010331'# (男)
        # Policy_number='9005010325'# 
        # Policy_number='9005010101'# 
        # Policy_number='9005010109'# 
        # Policy_number='9005010330'
        # Policy_number='9005010332'
        Policy_number='9005010519'


    ## ------登入---------
        test= LoginPage(driver, url, u"人壽保險核心業務系統")
        test.open()
        test.switch_frame('fraInterface')
        test.input_username(username)
        test.input_password(password)       
        test.click_submit()
        print('登入結束')
        ## ------移動到綜合查詢_保全查詢_得到保單受理日期---------
        Preser_Date=test.mouse_To_QueryPage_For_PreservationQuery(Policy_number)
        print("受理日期:"+Preser_Date)
        # ## ------修改營業日------------------------------------------------------
        test.mouse_To_SystemTime_for_preservation(Preser_Date)
        # ## ------契變撤銷作業----------------------------------------------------
        test.mouse_To_Preservation_rejectApply()
        test.input_PreservationPage_reject_apply(Policy_number)

        test.mouse_To_Preservation_rejectAgree()
        test.input_PreservationPage_reject_agree(Policy_number)
        driver.quit()

        driver2 = webdriver.Chrome(options=options)
        driver2.implicitly_wait(30)
        test2= LoginPage(driver2, url, u"人壽保險核心業務系統")
        test2.open()
        test2.switch_frame('fraInterface')
        test2.input_username(username)
        test2.input_password(password)       
        test2.click_submit()
        print('登入結束')


        # # ## ------查詢交費對應日，並移到前一個月的10號-------------------------------
        last_monDate=test2.mouse_To_QueryPage_For_RenewPremiumDate_for_Preservation(Policy_number)
        test2.mouse_To_SystemTime_for_preservation(last_monDate)
        # ## ------移動到保全處理---------
        test2.mouse_To_Preservation()
        
        test2.click_PreservationPage_ApplyBtn()
        test2.input_PreservationPage_PolicyNumber(Policy_number)
        test2.select_PreservationPage_AppType()

        test2.select_PreservationPage_TypeChange()
        # 4個簽名打勾
        test2.signup_PreservationPage()
        # 輸入完畢
        test2.submit_PreservationPage()

        # 自動核保
        test2.mouse_To_Preservation_Autocomplete(Policy_number)

        # 移動到契變審核作業
        test2.mouse_To_Preservation_Review()

        # 契變審核作業
        EdorNumber=test2.Preservation_confirm(Policy_number,last_monDate)

        driver2.quit()

        driver3 = webdriver.Chrome(options=options)
        driver3.implicitly_wait(30)
        test3= LoginPage(driver3, url, u"人壽保險核心業務系統")
        test3.open()
        test3.switch_frame('fraInterface')
        test3.input_username(username)
        test3.input_password(password)       
        test3.click_submit()
        print('登入結束')

        # 移動到綜合查詢_保全查詢_明細內的補退費訊息並拍照
        test3.mouse_To_QueryPage_For_PreserEdorNumber(EdorNumber)
        
        time.sleep(2)


        # ## ------移動到保全處理---------
        # test.mouse_To_Preservation()

        # # # 初次進行保全申請步驟
        # # test.click_PreservationPage_ApplyBtn()
        # # test.input_PreservationPage_PolicyNumber(Policy_number)
        # # test.select_PreservationPage_AppType()

        # # 已保全申請過的步驟
        # test.click_PreservationPage_PrivatePool(Policy_number)

        # # 選擇繳別變更
        # test.select_PreservationPage_TypeChange()
        # test.FillIn_PreservationPage_CRS()

    except Exception as e:
        print('有錯誤'+str(e))
        traceback.print_exc()
        driver.quit()

    finally:
        print('測試結束')
        driver.quit()



if __name__ == "__main__":
    pytest.main()