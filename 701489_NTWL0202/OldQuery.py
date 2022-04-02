import os
import sys
import traceback
import pytest
import time

o_path=os.path.abspath(os.path.join(os.getcwd(),'..'))
sys.path.append(o_path)

from U_LoginPage import LoginPage
from selenium import webdriver

def test_Query(my_fixture):
    try:
        personinfo=my_fixture[0]
        product=my_fixture[3]
        options = webdriver.ChromeOptions()
        options.add_argument('--ignore-certificate-errors')
        driver = webdriver.Chrome(options=options)
        # driver1 = webdriver.Chrome(options=options)
        # driver2 = webdriver.Chrome(options=options)
        # driver3 = webdriver.Chrome(options=options)
        # driver4 = webdriver.Chrome(options=options)
        driver.implicitly_wait(30)
        # driver1.implicitly_wait(30)
        # driver2.implicitly_wait(30)
        # driver3.implicitly_wait(30)
        # driver4.implicitly_wait(30)
        url ="https://10.1.113.23:9443/"
        username ="TEST41"
        password ="admin001"
        Acc_sub='1001002'
        # Date="2021-10-27"
        # Name="陳O應"

        # 原本的try:
    # ------登入---------
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
        Name=test.getName()
        
        print('保單號碼：'+Policy_number)
        print('客戶姓名：'+Name)
        test.mouse_To_button07()
        time.sleep(2)
        price=test.getRealPrice()
        print('保費'+price)

        # 跳回原本視窗
        test.switch_window_back()
        test.mouse_To_NewConform()
        # 新單覆核
        test.New_Conform(personinfo,Policy_number)
        time.sleep(2)
        driver.quit()

        # 重新開啟一個driver
        driver1 = webdriver.Chrome(options=options)
        driver1.implicitly_wait(30)
        test1= LoginPage(driver1, url, u"人壽保險核心業務系統")
        test1.open()
        test1.switch_frame('fraInterface')
        test1.input_username(username)
        test1.input_password(password)       
        test1.click_submit()
        print('登入結束')
        test1.mouse_To_Autoconfirm()
        test1.input_Policy_number(Policy_number)
        # 自動核保
        test1.click_Auto_Confirm_Search()
        test1.click_Auto_Confirm_select()
        test1.click_Auto_Confirm_Submit()
        driver1.quit()

        # 人工核保
        # 重開一個driver
        driver2 = webdriver.Chrome(options=options)
        driver2.implicitly_wait(30)
        test2= LoginPage(driver2, url, u"人壽保險核心業務系統")
        test2.open()
        test2.switch_frame('fraInterface')
        test2.input_username(username)
        test2.input_password(password)       
        test2.click_submit()
        print('登入結束')
        # 移動到人工核保
        test2.mouse_To_Articalconfirm()

        # 輸入公共池內的保單號碼
        test2.input_ArtiConfirmPage_first_PolicyNumber(Policy_number)

        # 點擊公共池內的查詢
        test2.click_ArtiConfirmPage_first_QBtn()

        # 點擊公共池內的保單
        test2.click_ArtiConfirmPage_PublicPool()

        # 輸入個人池內的保單號碼
        # test2.input_ArtiConfirmPage_second_PoliceNumber(Policy_number)

        # 點擊個人池內的查詢
        # test2.click_ArtiConfirmPage_second_Qbtn()

        # 點擊個人池內的保單
        # test2.click_ArtiConfirmPage_second_PrivatePool()

        # 全選照會
        test2.click_ArtiConfirmPage_checkAll()

        # 解除按鈕
        test2.click_ArtiConfirmPage_AutoUWButton()

        # 選擇核保結論方法
        test2.sele_ArtiConfirmPage_auwState()

        #選擇核保流向方法
        test2.sele_ArtiConfirmPage_uwUpReport()

        # 輸入意見
        test2.input_ArtiConfirmPage_UWIdea()

        # 點擊確定按鈕
        test2.click_ArtiConfirmPage_button1()
        driver2.quit()
        

        # 立帳
        driver3 = webdriver.Chrome(options=options)
        driver3.implicitly_wait(30)
        test3= LoginPage(driver3, url, u"人壽保險核心業務系統")
        test3.open()
        test3.switch_frame('fraInterface')
        test3.input_username(username)
        test3.input_password(password)       
        test3.click_submit()
        print('登入結束')
        # test3.mouse_To_AccountQuery()
        # ASN=test3.query_AccountPage_serial_number(product)
        # print('立帳流水號：'+ASN)
    # # -------單筆立帳-------------------------
        test3.mouse_To_Account()
        # 會計科目
        test3.select_Account_subject(Acc_sub)
        # 金額
        test3.input_Account_Money(price)
        # 摘要(保單號碼)
        test3.input_Account_Policy(Policy_number)
        # 新增
        test3.click_Account_add_btn()
        # 立帳
        test3.click_Account_setup_btn()
        # -------查詢立帳流水號--------------------
        test3.mouse_To_AccountQuery()
        ASN=test3.query_AccountPage_serial_number(price)
        print('立帳流水號：'+ASN)
    # -------單筆自繳建檔--------------------
        test3.mouse_To_Premium()
        # 選擇單據類型
        test3.select_PremiumPage_DocumentType_cash()
        # 點擊新建按鈕
        test3.click_PremiumPage_AddBtn()
    # -------轉換視窗進行單據建檔--------------------
        handles = driver3.window_handles
        # print(len(handles))
        driver3.switch_to.window(handles[1])
        t0=driver3.title
        print('目前窗口：'+t0)
        # 輸入單據金額
        test3.input_PremiumPage_DocumentMoney(price)
        # 輸入使用金額
        test3.input_PremiumPage_UseMoney(price)
        # 輸入立帳流水號
        test3.input_PremiumPage_AccountNo(ASN)
        # 輸入繳款日期
        test3.input_PremiumPage_EnteraccDate(personinfo)
        # 點擊保存按鈕
        test3.click_PremiumPage_savebtn()
        # 選擇單據訊息結果的第一個選項
        test3.click_PremiumPage_SelectResult()
        # 選擇繳費類型
        test3.select_PremiumPage_PayType()
        # 選擇首期保費並全部提交
        test3.select_PremiumPage_BusinessType(price,Policy_number)
        driver3.quit()

        # 單筆自繳審核
        driver4 = webdriver.Chrome(options=options)
        driver4.implicitly_wait(30)
        test4= LoginPage(driver4, url, u"人壽保險核心業務系統")
        test4.open()
        test4.switch_frame('fraInterface')
        test4.input_username(username)
        test4.input_password(password)       
        test4.click_submit()
        print('登入結束')
        test4.mouse_To_PremiumCheck()
        # 輸入金額範圍上限
        test4.input_PremiumCheckPage_Amount(price)
        # 輸入金額範圍下限
        test4.input_PremiumCheckPage_Amount2(price)
        # 選擇幣別
        test4.select_PremiumCheckPage_Currency()
        # 選擇繳費方式
        test4.select_PremiumCheckPage_DocumentType()
        # 輸入銀行入帳起日
        test4.input_PremiumCheckPage_StartPayDate(personinfo)
        # 輸入銀行入帳迄日
        test4.input_PremiumCheckPage_EndPayDate(personinfo)
        # 點擊查詢按鈕
        test4.click_PremiumCheckPage_Qbtn()
        # 點擊水池內第一個選項
        test4.click_PremiumCheckPage_PoolFirst()
        # 點擊全部提交
        test4.submit_PremiumCheckPage_All()

        # driver.close()
        # driver.quit()
        # driver1.quit()
        # driver2.quit()
        # driver3.quit()
        # driver4.quit()

    except Exception as e:
        print('有錯誤'+str(e))
        traceback.print_exc()
        # driver.close()
        driver.quit()
        driver1.quit()
        driver2.quit()
    finally:
        print('測試結束')
        driver.quit()
        driver1.quit()
        driver2.quit()
        driver3.quit()

if __name__ == "__main__":
    pytest.main()
