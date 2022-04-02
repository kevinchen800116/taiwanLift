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
    personinfo=my_fixture[0]
    SaleReport=my_fixture[1]
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

        
        # 依照幣別判斷立帳金額為台幣或外幣
        if(SaleReport["Currency"]=="NTD"):
            # ## 會計科目：現金
            Acc_sub='1001002'
        elif(SaleReport["Currency"]=="USD"):
            ## 會計科目：USD
            Acc_sub='149000106'

        test= LoginPage(driver, url, u"人壽保險核心業務系統")
        test.open()
        test.switch_frame('fraInterface')
        test.input_username(username)
        test.input_password(password)       
        test.click_submit()
        print('登入結束')

        ## 更改系統日期至7/1
        test.mouse_To_SystemTime(personinfo)
    ## ------查詢保單號碼---------
        test.mouse_To_NewQuery()

        # test.input_QueryPage_Date(personinfo)
        # test.input_QueryPage_Name(personinfo)
        # test.input_QueryPage_ID(personinfo)
        test.input_QueryPage_PolicyNB(personinfo)
        test.click_QueryPage_Button()
        test.check_QueryPage_seleOne()

        ## 取得保單號碼
        Policy_number=test.getNumbertext()
        # 取得要保人姓名
        Name=test.getName()
        
        print('保單號碼:'+Policy_number)
        print('客戶姓名:'+Name)

        ## 查詢繳費資訊
        test.mouse_To_button07()
        # test.take_screenshot()

        time.sleep(2)
        price=test.getRealPrice()
        print('保費'+price)

        print("首期自行繳費，續期銀轉")
        ## ----------立帳--------------------
        driver3 = webdriver.Chrome(options=options)
        driver3.implicitly_wait(30)
        test3= LoginPage(driver3, url, u"人壽保險核心業務系統")
        test3.open()
        test3.switch_frame('fraInterface')
        test3.input_username(username)
        test3.input_password(password)       
        test3.click_submit()
        print('登入結束')

        # 更改系統日期
        test3.mouse_To_SystemTime(personinfo)

        ## -------單筆立帳-------------------------
        test3.mouse_To_Account()
        ## 會計科目
        test3.select_Account_subject(Acc_sub)
        ## 金額
        test3.input_Account_Money(price)
        ## 摘要(保單號碼)
        test3.input_Account_Policy(Policy_number)
        ## 新增
        test3.click_Account_add_btn()
        ## 立帳
        test3.click_Account_setup_btn()

        ## -------單筆自繳建檔--------------------
        test3.mouse_To_Premium()
        ## 選擇單據類型
        test3.select_PremiumPage_DocumentType_cash(Acc_sub)
        ## 點擊新建按鈕
        test3.click_PremiumPage_AddBtn()
        ## -------轉換視窗進行單據建檔--------------------
        handles = driver3.window_handles
        ## print(len(handles))
        driver3.switch_to.window(handles[1])
        t0=driver3.title
        print('目前窗口:'+t0)
        ## 輸入單據金額(匯款)
        test3.input_PremiumPage_DocumentMoney3(price)


        ## 輸入繳款日期
        test3.input_PremiumPage_EnteraccDate3(personinfo)
        
        # 如果為外幣會進來
        if(Acc_sub=='149000106'):
            # 輸入銀行代號
            test3.input_PremiumPage_BankCode()

            # 輸入銀行帳號 & 輸入繳款人帳號
            test3.input_PremiumPage_BankAcc()

        ## 點擊保存按鈕
        test3.click_PremiumPage_savebtn()
        ## 選擇單據訊息結果的第一個選項
        test3.click_PremiumPage_SelectResult()
        ## 選擇繳費類型
        test3.select_PremiumPage_PayType()
        ## 選擇首期保費並全部提交
        print("提交的保費:"+price)
        time.sleep(2)
        test3.select_PremiumPage_BusinessType3(price,Policy_number)

        driver3.quit()

        ## ----------簽發以及簽收保單-------------------------------------
        driver5 = webdriver.Chrome(options=options)
        driver5.implicitly_wait(30)
        test5= LoginPage(driver5, url, u"人壽保險核心業務系統")
        test5.open()
        test5.switch_frame('fraInterface')
        test5.input_username(username)
        test5.input_password(password)       
        test5.click_submit()
        print('登入結束')

        ## ------移動到簽發保單---------
        # 設定系統時間為7/1 
        test5.mouse_To_SystemTime(personinfo)
        test5.mouse_To_SignPage()
        ## 輸入保單號碼
        test5.input_SignPage_policyNumbr(Policy_number)
        ## 點擊查詢按鈕
        test5.click_SignPage_Qbtn()
        ## 選擇	公共工作池保單
        test5.click_SignPage_checksele()
        ## 發送簽發保單
        test5.click_SignPage_signbutton()
        time.sleep(1)

        ## ------移動到列印保單---------
        test5.mouse_To_PrintSignPage()
        ## 輸入保單號碼
        test5.input_PrintSignPage_PolicyNumber(Policy_number)
        # 點擊查詢按鈕
        test5.click_PrintSignPage_Qbtn()
        ## 選擇池內第一個選項
        test5.click_PrintSignPage_PoolFirst()
        time.sleep(1)
        ##點擊列印保單
        test5.click_PrintSignPage_PrintBtn()
        time.sleep(1)

        ## ------移動到簽收保單---------
        test5.mouse_To_SignBackPage()
        test5.select_SignBackPage_agreeBtn()
        test5.input_SignBackPage_Policy_number(Policy_number)
        test5.input_SignBackPage_SignBackDate(personinfo)
        test5.click_SignBackPage_checkBtn()
        test5.click_SignBackPage_submitBtn()
        time.sleep(1)
        driver5.quit()


        # ### 續期銀行轉帳(授權書建檔)
        # BankTransfer(Policy_number,personinfo)
        
        # ### 續期銀轉 (跑批次得到續期繳費通知書) (需手動製磁)
        # Query_RenewPremiumDate(Policy_number)
        

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
