
import time
import traceback
import os
import sys

### 移動到外層目錄(ZREALtest)是為了取得U_LoginPage這個py檔內的LoginPage class
o_path=os.path.abspath(os.path.join(os.getcwd(),'..'))
sys.path.append(o_path)

from U_LoginPage import LoginPage
from selenium import webdriver

# 新單覆核 &自動核保 &人工核保 &首期繳費 &簽發保單 &列印保單 &簽收保單 
def Query_Confirm_forTest(my_fixture):
    try:
        personinfo=my_fixture[0]
        SaleReport=my_fixture[1]
        options = webdriver.ChromeOptions()
        options.add_argument('--ignore-certificate-errors')
        # options.add_argument('--headless') #  無頭
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
        elif(SaleReport["Currency"]=="CNY"):
          Acc_sub='149000114'
        elif(SaleReport["Currency"]=="AUD"):
          Acc_sub='149000109'
        elif(SaleReport["Currency"]=="EUR"):
          Acc_sub='149000922'
        elif(SaleReport["Currency"]=="ZAR"):
          Acc_sub='149000A02'
    ## ------登入---------
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
        test.input_QueryPage_Date(personinfo)
        test.input_QueryPage_Name(personinfo)
        test.input_QueryPage_ID(personinfo)
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

        ## 跳回原本視窗
        test.switch_window_back()
        test.mouse_To_NewConform()
        # 新單覆核
        test.New_Conform(personinfo,Policy_number)
        time.sleep(2)
        driver.quit()

        ## 重新開啟一個driver
        driver1 = webdriver.Chrome(options=options)
        driver1.implicitly_wait(30)
        test1= LoginPage(driver1, url, u"人壽保險核心業務系統")
        test1.open()
        test1.switch_frame('fraInterface')
        test1.input_username(username)
        test1.input_password(password)       
        test1.click_submit()
        print('登入結束')
        # 更改系統日期
        test1.mouse_To_SystemTime(personinfo)
        ## 移動到自動核保
        test1.mouse_To_Autoconfirm()
        test1.input_Policy_number(Policy_number)
        time.sleep(2)
        ## 自動核保
        test1.click_Auto_Confirm_Search()
        test1.click_Auto_Confirm_select()
        test1.click_Auto_Confirm_Submit()
        time.sleep(2)
        driver1.quit()

        ## 人工核保
        ## 重開一個driver
        driver2 = webdriver.Chrome(options=options)
        driver2.implicitly_wait(30)
        test2= LoginPage(driver2, url, u"人壽保險核心業務系統")
        test2.open()
        test2.switch_frame('fraInterface')
        test2.input_username(username)
        test2.input_password(password)       
        test2.click_submit()
        print('登入結束')
        ## 更改系統日期
        test2.mouse_To_SystemTime(personinfo)
        ## 移動到人工核保
        test2.mouse_To_Articalconfirm()

        ## 輸入公共池內的保單號碼
        test2.input_ArtiConfirmPage_first_PolicyNumber(Policy_number)

        ## 點擊公共池內的查詢
        test2.click_ArtiConfirmPage_first_QBtn()

        ## 點擊公共池內的保單
        test2.click_ArtiConfirmPage_PublicPool()

        ## 輸入個人池內的保單號碼
        # test2.input_ArtiConfirmPage_second_PoliceNumber(Policy_number)

        ## 點擊個人池內的查詢
        # test2.click_ArtiConfirmPage_second_Qbtn()

        ## 點擊個人池內的保單
        # test2.click_ArtiConfirmPage_second_PrivatePool()

        ## 全選照會(新契約人工核保)
        test2.click_ArtiConfirmPage_checkAll()

        ## 解除按鈕
        test2.click_ArtiConfirmPage_AutoUWButton()

        ## 選擇核保結論方法
        test2.sele_ArtiConfirmPage_auwState()

        ## 選擇核保流向方法
        test2.sele_ArtiConfirmPage_uwUpReport()

        ## 輸入意見
        test2.input_ArtiConfirmPage_UWIdea()

        ## 點擊確定按鈕
        test2.click_ArtiConfirmPage_button1()
        driver2.quit()

    ### 如果為首期(PayMode2=1)才會立帳繳費，續期需自行手動。
    ### ------首期自行繳費 (續期自行繳款還沒做)------------------------------------------
        if(personinfo["PayMode2"]=="1" and personinfo["SecPayMode2"] == "1"):
            print("進入首期自行繳費")
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



          #     ## -------查詢立帳流水號(目前被拿掉不需要查詢了)--------------------
          #     test3.mouse_To_AccountQuery()
          #     ASN=test3.query_AccountPage_serial_number(price,Acc_sub)
          #     print('立帳流水號：'+ASN)
        
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

            if(SaleReport["Currency"]=="NTD"):
              ## 輸入單據金額(現金)
              test3.input_PremiumPage_DocumentMoney(price)
              ## 輸入繳款日期(現金)
              test3.input_PremiumPage_EnteraccDate(personinfo)
            else:
              ## 輸入單據金額(匯款)
              test3.input_PremiumPage_DocumentMoney3(price)
              ## 輸入繳款日期(匯款)
              test3.input_PremiumPage_EnteraccDate3(personinfo)
            # elif(SaleReport["Currency"]=="USD"):
            #   ## 輸入單據金額(匯款)
            #   test3.input_PremiumPage_DocumentMoney3(price)
            #   ## 輸入繳款日期(匯款)
            #   test3.input_PremiumPage_EnteraccDate3(personinfo)
            
            # 如果為外幣會進來
            # if(Acc_sub=='149000106'):
            if(Acc_sub !='1001002'):
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

          # ## ----------單筆自繳審核(目前被拿掉)-------------------------------------
            # driver4 = webdriver.Chrome(options=options)
            # driver4.implicitly_wait(30)
            # test4= LoginPage(driver4, url, u"人壽保險核心業務系統")
            # test4.open()
            # test4.switch_frame('fraInterface')
            # test4.input_username(username)
            # test4.input_password(password)       
            # test4.click_submit()
            # print('登入結束')
            # ## 更改系統日期
            # test4.mouse_To_SystemTime(personinfo)
            # test4.mouse_To_PremiumCheck()
            # ## 輸入金額範圍上限
            # test4.input_PremiumCheckPage_Amount(price)
            # ## 輸入金額範圍下限
            # test4.input_PremiumCheckPage_Amount2(price)
            # ## 選擇幣別
            # test4.select_PremiumCheckPage_Currency()
            # ## 選擇繳費方式
            # test4.select_PremiumCheckPage_DocumentType()
            # ## 輸入銀行入帳起日
            # test4.input_PremiumCheckPage_StartPayDate(personinfo)
            # ## 輸入銀行入帳迄日
            # test4.input_PremiumCheckPage_EndPayDate(personinfo)
            # ## 點擊查詢按鈕
            # test4.click_PremiumCheckPage_Qbtn()
            # ## 點擊水池內第一個選項
            # test4.click_PremiumCheckPage_PoolFirst()
            # ## 點擊全部提交
            # success=test4.submit_PremiumCheckPage_All()
            # assert success == "操作成功。"
            # print(success+"繳費完成~")
            # driver4.quit()

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
            # success=test5.click_SignBackPage_submitBtn()
            time.sleep(1)

          #   ## ------移動到批次處理任務執行(需增加投資型或傳統型的判斷)---------
            if(SaleReport["MainPolNo"]=='NITU0905'):
              test5.mouse_To_SystemTaskSetup()
              test5.input_TaskPage_BussNo(Policy_number)
              test5.input_TaskPage_PolicyNo(Policy_number)

              test5.input_TaskPage_TaskCode('NB001')
              test5.input_TaskPage_ExeDate(personinfo,'NB001')
              test5.click_TaskPage_executeTask()

              test5.mouse_To_SystemTaskSetup()
              test5.input_TaskPage_BussNo(Policy_number)
              test5.input_TaskPage_PolicyNo(Policy_number)

              test5.input_TaskPage_TaskCode('ILP015')
              test5.input_TaskPage_ExeDate(personinfo,'ILP015')
              test5.click_TaskPage_executeTask()
              # Task_success=test5.verify_TaskPage_executeDone()
              time.sleep(2)
            driver5.quit()

    ### ------首期自行繳費 續期銀轉(續期銀轉的跑批次還沒做) ------------------------------------------
        elif(personinfo["PayMode2"]=="1" and personinfo["SecPayMode2"] == "2"):
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

            if(SaleReport["Currency"]=="NTD"):
              ## 輸入單據金額(現金)
              test3.input_PremiumPage_DocumentMoney(price)
              ## 輸入繳款日期(現金)
              test3.input_PremiumPage_EnteraccDate(personinfo)
            elif(SaleReport["Currency"]=="USD"):
              ## 輸入單據金額(匯款)
              test3.input_PremiumPage_DocumentMoney3(price)
              ## 輸入繳款日期(匯款)
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
            test5.take_screenshot()
            time.sleep(1)
            driver5.quit()
            

            # return Policy_number

          #   ## ------移動到批次處理任務執行(需增加投資型或傳統型的判斷)---------
            # test5.mouse_To_SystemTaskSetup()
            # test5.input_TaskPage_BussNo(Policy_number)
            # test5.input_TaskPage_PolicyNo(Policy_number)

            # test5.input_TaskPage_TaskCode('NB001')
            # test5.input_TaskPage_ExeDate(personinfo,'NB001')
            # test5.click_TaskPage_executeTask()

            # test5.mouse_To_SystemTaskSetup()
            # test5.input_TaskPage_BussNo(Policy_number)
            # test5.input_TaskPage_PolicyNo(Policy_number)

            # test5.input_TaskPage_TaskCode('ILP015')
            # test5.input_TaskPage_ExeDate(personinfo,'ILP015')
            # test5.click_TaskPage_executeTask()
            # Task_success=test5.verify_TaskPage_executeDone()
            # time.sleep(2)

            return Policy_number
            # ### 續期銀行轉帳(授權書建檔)
            # BankTransfer(Policy_number,personinfo)
            
            # ### 續期銀轉 (跑批次得到續期繳費通知書) (需手動製磁)
            # Query_RenewPremiumDate(Policy_number,personinfo)

    ### ------首期銀轉繳費(續期自行繳款還沒做) ------------------------------------------
        elif(personinfo["PayMode2"]=="2" and personinfo["SecPayMode2"] == "1"):
            print("進入首期銀轉 繳費，請手動繳")

            return Policy_number


    ### ------首期/續期皆銀轉繳費 (續期銀轉的跑批次還沒做) ------------------------------------------
        elif(personinfo["PayMode2"]=="2" and personinfo["SecPayMode2"] == "2"):
            print("進入首/續期皆銀轉繳費，請手動繳")

            return Policy_number
            



    except Exception as e:
        traceback.print_exc()
        print('有錯誤'+str(e))
        driver.quit()
        # driver1.quit()
        # driver2.quit()
        # driver3.quit()
        # driver4.quit()
        # driver5.quit()
    finally:
      # if(personinfo["PayMode2"]=="1"):
      #     assert Task_success == "批处理执行完成!"
      #     print(Task_success+'測試結束')
        driver.quit()
        # driver1.quit()
        # driver2.quit()
        # driver3.quit()
        # driver4.quit()

# 綜合查詢-->查詢續期繳費日 & 執行CDPA01產生"續期繳費通知書"
def Query_RenewPremiumDate(Policy_number,personinfo,RenewPayDate):
    # personinfo=my_fixture[0]
    try:
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

     # ### --------綜合查詢_續期交費日-----------------------------------------------------------
       # ###  查詢交費對應日
        renewable_premium_Date = test.mouse_To_QueryPage_For_RenewPremiumDate(Policy_number)
       # ### --------續期批次作業CDPA01得到續期繳費通知書------------------------
        test.mouse_To_SystemTaskSetup()
        for i, r in zip(Policy_number,RenewPayDate):
          test.input_TaskPage_PolicyNo(i)
          test.input_TaskPage_BussNo(i)
          test.input_TaskPage_TaskCode("CDPA01")
          # test.input_TaskPage_RAW_ExeDate(personinfo)
          test.input_TaskPage_RAW_ExeDate(r)
          test.click_TaskPage_executeTask()

        # ### --------移動到通知書查詢---------
        test.mouse_To_Renew()
        for i in range(len(Policy_number)):
          test.input_Renew_PolicyNumber(Policy_number)
          time.sleep(5)
          test.click_Renew_Qbtn()
          ### 得到續期繳費通知書的催告日
          NoticeDate=test.get_Renew_NoticeDate()
          ### 得到此日期是為了執行CDPA03(此處未寫return所以沒辦法把NoticeDate傳出去，之後要補) 
          print("續期繳費通知書的催告日:"+NoticeDate)
          test.take_screenshot()
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

# (銀轉)首/續期  若為續期應先執行Query_RenewPremiumDate()，查詢續期繳費日期，並執行批次CDPA01
def BankTransfer(Policy_number,personinfo):
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


        # driver = webdriver.Chrome(options=options)
        # driver.implicitly_wait(30)
        url ="https://10.1.113.23:9443/"
        username ="TEST41"
        password ="admin001"
        # # ---------------移動到銀行轉帳->授權書建檔---------------
        # test= LoginPage(driver, url, u"人壽保險核心業務系統")
        # test.open()
        # test.switch_frame('fraInterface')
        # test.input_username(username)
        # test.input_password(password)       
        # test.click_submit()
        # print('登入結束')
        # # 銀行轉帳

        # test.mouse_To_BankTransfer_AuthBook()

        # # 點擊新增授權書按鈕
        # test.click_BankAZ_NewAuthorBook_Btn()

        # # 輸入授權書條碼
        # test.input_BankAZ_AuthorBar(Policy_number)
        
        # # 點擊授權書訊息內的按鈕從左至右
        # test.click_BankAZ_SaveBTN1()
        # # 點擊授權書詳細訊息內的按鈕從左至右
        # test.click_BankAZ_SaveBTN2()
        # driver.quit()


        # # # --------移動到銀轉_送件---------------
        driver1 = webdriver.Chrome(options=options)
        driver1.implicitly_wait(30)
        test1= LoginPage(driver1, url, u"人壽保險核心業務系統")
        test1.open()
        test1.switch_frame('fraInterface')
        test1.input_username(username)
        test1.input_password(password)       
        test1.click_submit()
        print('登入結束')


        test1.mouse_To_BankTransfer_Delivery()
        test1.input_BD_AuthorBar1(Policy_number)
        test1.click_BD_CheckBtn()
        test1.click_BD_DeliveryBtn()
        time.sleep(5)
        driver1.quit()

        # # --------移動到銀轉_送核--------
        driver2 = webdriver.Chrome(options=options)
        driver2.implicitly_wait(30)
        test2= LoginPage(driver2, url, u"人壽保險核心業務系統")
        test2.open()
        test2.switch_frame('fraInterface')
        test2.input_username(username)
        test2.input_password(password)       
        test2.click_submit()
        print('登入結束')
        # # --------銀轉_送核--------
        test2.mouse_To_BankTransfer_Approval()
        test2.input_BA_AuthorBar1(Policy_number)
        test2.click_BA_checkBtn()
        test2.click_BA_ApproveBtn()

        # # -------銀轉_回件作業--------
        test2.mouse_To_BankTransfer_Return()
        if(personinfo["bankcode"]=="1080014" or personinfo["bankcode"]=="8080015"):
          # 5.使用send_keys方法上傳檔案
          test2.upload_BF_file(personinfo)
          test2.click_BF_saveBtn(personinfo)
          test2.take_screenshot()
        else:
          test2.select_BR_result()
          test2.input_BR_AuthorBar(Policy_number)
          test2.click_BF_submitBtn()
          test2.take_screenshot()
          time.sleep(1)

        driver2.quit()

    except Exception as e:
        print('有錯誤'+str(e))
        traceback.print_exc()
        # driver.close()
        # driver.quit()
        driver1.quit()
        driver2.quit()
    finally:
        # assert success == "操作成功。"
        # print(success)
        print('銀轉測試結束')
        # driver.quit()
        driver1.quit()
        driver2.quit()

