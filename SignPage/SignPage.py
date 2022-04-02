import os
import sys
import traceback
import pytest
import time

o_path=os.path.abspath(os.path.join(os.getcwd(),'..'))
sys.path.append(o_path)

from U_LoginPage import LoginPage
from selenium import webdriver


def test_SignPage(my_fixture):
    Policy_number=my_fixture
    print("我是"+str(Policy_number))
    try:
        options = webdriver.ChromeOptions()
        options.add_argument('--ignore-certificate-errors')
        options.add_experimental_option('prefs', {
            # 注意下载路径，wins下必须是 \  而不是 /
            "download.default_directory": r"D:\Users\701489\Desktop\PDF", #Change default directory for downloads
            "download.prompt_for_download": False, #To auto download the file 禁止弹出下載確認窗口，直接下載。
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True, #It will not show PDF directly in chrome
            "safebrowsing.enabled" : True # 說明： “safebrowsing.enabled”: True參數。增加了這個參數。就不會彈出保存與放棄的提示；
            })
        driver = webdriver.Chrome(options=options)
        driver.implicitly_wait(30)
        url ="https://10.1.113.23:9443/"
        username ="TEST41"
        password ="admin001"
        Policy_number='9000162722' #(首/續期)
        # Policy_number='9000143535' # (保全測試)
        # date_str='2021-08-20'
        # Policy_number='9000150233'# (承保列印測試)
        # Policy_number='9000151627'# (簽發保單測試)
        # personinfo={"AppntIDNo":"K182659514"}

        test= LoginPage(driver, url, u"人壽保險核心業務系統")
        test.open()
        test.switch_frame('fraInterface')
        test.input_username(username)
        test.input_password(password)       
        test.click_submit()
        print('登入結束')

# # ### --------綜合查詢_續期交費日-----------------------------------------------------------
#         renewable_premium_Date=test.mouse_To_QueryPage_For_RenewPremiumDate(Policy_number)
#        # ### --------續期批次作業CDPA01得到續期繳費通知書------------------------
#         test.mouse_To_SystemTaskSetup()
#         test.input_TaskPage_PolicyNo(Policy_number)
#         test.input_TaskPage_BussNo(Policy_number)
#         test.input_TaskPage_TaskCode("CDPA01")
#         test.input_TaskPage_RAW_ExeDate(renewable_premium_Date)
#         test.click_TaskPage_executeTask()

#         # ### --------移動到通知書查詢---------
#         test.mouse_To_Renew()
#         test.input_Renew_PolicyNumber(Policy_number)
#         test.click_Renew_Qbtn()
#         ### 得到續期繳費通知書的催告日
#         NoticeDate=test.get_Renew_NoticeDate()
#         print("續期繳費通知書的催告日:"+NoticeDate)

        # # ### --------批次作業CDPA03得到催告通知書--------
        # test.mouse_To_SystemTaskSetup()
        # test.input_TaskPage_PolicyNo(Policy_number)
        # test.input_TaskPage_BussNo(Policy_number)
        # test.input_TaskPage_TaskCode("CDPA03")
        # test.input_TaskPage_RAW_ExeDate(NoticeDate)
        # test.click_TaskPage_executeTask()

        # # ### --------移動到通知書查詢---------
        # test.mouse_To_Renew()
        # test.input_Renew_PolicyNumber(Policy_number)
        # test.click_Renew_Qbtn()
        # ### 得到催告通知書的自動墊繳起日
        # APD=test.get_Renew_APD()
        # print("催告通知書的自動墊繳起日:"+APD)
    # ### ---------------停效--------------------------------------------------------------
    #     # ### --------批次作業CDPA04得到停效通知書---------
    #     test.mouse_To_SystemTaskSetup()
    #     test.input_TaskPage_PolicyNo(Policy_number)
    #     test.input_TaskPage_BussNo(Policy_number)
    #     test.input_TaskPage_TaskCode("CDPA04")
    #     test.input_TaskPage_RAW_ExeDate(APD)
    #     test.click_TaskPage_executeTask()

    #     # ### --------批次作業CDPA06得到停效通知書---------
    #     test.mouse_To_SystemTaskSetup()
    #     test.input_TaskPage_PolicyNo(Policy_number)
    #     test.input_TaskPage_BussNo(Policy_number)
    #     test.input_TaskPage_TaskCode("CDPA06")

    ##    自行將自動墊繳日+35天後(2021-12-21) 執行CDPA06
    #     test.input_TaskPage_ExeDate_Plus35(APD)
    #     test.click_TaskPage_executeTask()
    # ### ---------------------------------------------------------------------------------
### ---------------------------------------------------------------------------------
       # ### --------移動到理賠案件---------
        # test.mouse_To_claim()
        # test.click_ClaimPage_newApplyBtn()
        # test.input_ClaimPage_IdNo(personinfo)
        # test.select_ClaimPage_RgtSourcesType()
        # test.input_ClaimPage_AcceptedDate()
        # test.input_ClaimPage_AccidentDate()
        # test.select_ClaimPage_occurReason()
        # test.click_ClaimPage_AccReason()

        ### ------移動系統管理---------
        ## 設定系統時間為7/1 
        # test.mouse_To_SystemTime_JulyOne()


        ### ------移動到列印保單---------
        # test.mouse_To_PrintSignPage()
        # ## 輸入保單號碼
        # test.input_PrintSignPage_PolicyNumber(Policy_number)
        # # 點擊查詢按鈕
        # test.click_PrintSignPage_Qbtn()
        # ## 選擇池內第一個選項
        # test.click_PrintSignPage_PoolFirst()
        # ##點擊列印保單
        # test.click_PrintSignPage_PrintBtn()

        ### ------移動到保全處理---------
        # test.mouse_To_Preservation()

        ## 初次進行保全申請步驟
        # test.click_PreservationPage_ApplyBtn()
        # test.input_PreservationPage_PolicyNumber(Policy_number)
        # test.select_PreservationPage_AppType()

        # # 已保全申請過的步驟
        # test.click_PreservationPage_PrivatePool(Policy_number)

        # # 選擇繳別變更
        # test.select_PreservationPage_TypeChange()
        # test.FillIn_PreservationPage_CRS()


        ## ------移動到簽發保單---------
        # 設定系統時間為7/1 
        # test.mouse_To_SystemTime_JulyOne()
        # test.mouse_To_SignPage()
        # ## 輸入保單號碼
        # test.input_SignPage_policyNumbr(Policy_number)
        # ## 點擊查詢按鈕
        # test.click_SignPage_Qbtn()
        # ## 選擇	公共工作池保單
        # test.click_SignPage_checksele()
        # ## 發送簽發保單
        # test.click_SignPage_signbutton()

        # time.sleep(1)

        # ## ------移動到列印保單---------
        # test.mouse_To_PrintSignPage()
        # ## 輸入保單號碼
        # test.input_PrintSignPage_PolicyNumber(Policy_number)
        # # 點擊查詢按鈕
        # test.click_PrintSignPage_Qbtn()
        # ## 選擇池內第一個選項
        # test.click_PrintSignPage_PoolFirst()
        # time.sleep(1)
        # ##點擊列印保單
        # test.click_PrintSignPage_PrintBtn()
        # time.sleep(1)

        ## ------移動到簽收保單---------
        # test.mouse_To_SignBackPage()
        # test.select_SignBackPage_agreeBtn()
        # test.input_SignBackPage_Policy_number(Policy_number)
        # test.input_SignBackPage_SignBackDate(Policy_number)
        # test.click_SignBackPage_checkBtn()
        # test.click_SignBackPage_submitBtn()
        # time.sleep(1)

        # 移動到批次處理任務執行
        # test.mouse_To_SystemTaskSetup()
        # test.input_TaskPage_BussNo(Policy_number)
        # test.input_TaskPage_PolicyNo(Policy_number)

        # test.input_TaskPage_TaskCode('NB001')
        # test.input_TaskPage_ExeDate(Policy_number,'NB001')
        # test.click_TaskPage_executeTask()

        # test.input_TaskPage_TaskCode('ILP015')
        # test.input_TaskPage_ExeDate(Policy_number,'ILP015')
        # test.click_TaskPage_executeTask()
        # time.sleep(1)
    # # --------移動到銀轉_授權書建檔--------
        # test.mouse_To_BankTransfer_AuthBook()
        # # 點擊新增授權書按鈕
        # test.click_BankAZ_NewAuthorBook_Btn()
        # test.input_BankAZ_AuthorBar(Policy_number)
        # test.click_BankAZ_SaveBTN1()
        # test.click_BankAZ_SaveBTN2()
    # # # --------移動到銀轉_送件---------------
        # test.mouse_To_BankTransfer_Delivery()
        # test.input_BD_AuthorBar1(Policy_number)
        # test.click_BD_CheckBtn()
        # test.click_BD_DeliveryBtn()

    # # --------移動到銀轉_送核--------
        # test.mouse_To_BankTransfer_Approval()
        # test.input_BA_AuthorBar1(Policy_number)
        # test.click_BA_checkBtn()
        # test.click_BA_ApproveBtn()

        test.mouse_To_BankTransfer_Return()
        test.select_BR_result()
        test.input_BR_AuthorBar(Policy_number)
        test.click_BF_submitBtn()

        time.sleep(10)
        driver.quit()  
        

    except Exception as e:
        print('有錯誤'+str(e))
        traceback.print_exc()
        # driver.close()
        driver.quit()
    finally:
        # assert success == "操作成功。"
        # print(success)
        print('測試結束')
        driver.quit()

if __name__ == "__main__":
    pytest.main()
