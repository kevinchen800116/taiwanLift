import os
import sys
import traceback
import pytest
import time

o_path=os.path.abspath(os.path.join(os.getcwd(),'..'))
sys.path.append(o_path)

from U_LoginPage import LoginPage
from selenium import webdriver


def test_PremiumAdd(my_fixture):
    try:
        personinfo=my_fixture[0]
        product=my_fixture[3]
        options = webdriver.ChromeOptions()
        options.add_argument('--ignore-certificate-errors')
        driver = webdriver.Chrome(options=options)
        driver.implicitly_wait(30)
        url ="https://10.1.113.23:9443/"
        username ="TEST41"
        password ="admin001"

        # 立帳科目：庫存現金
        Acc_sub='1001002'

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
        test.input_QueryPage_ID(personinfo)
        
        Policy_number=test.getNumbertext()
        Name=test.getName()

        print('保單號碼：'+Policy_number)
        print('客戶姓名：'+Name)
    # -------查詢立帳流水號--------------------
        test.mouse_To_AccountQuery()
        ASN=test.query_AccountPage_serial_number(product)
        print('立帳流水號：'+ASN)
    # # -------單筆立帳-------------------------
        test.mouse_To_Account()
        # 會計科目
        test.select_Account_subject(Acc_sub)
        # 金額
        test.input_Account_Money(product)
        # 摘要(保單號碼)
        test.input_Account_Policy(Policy_number)
        # 新增
        test.click_Account_add_btn()
        # 立帳
        test.click_Account_setup_btn()
    # -------單筆自繳建檔--------------------
        test.mouse_To_Premium()
        # 選擇單據類型
        test.select_PremiumPage_DocumentType_cash()
        # 點擊新建按鈕
        test.click_PremiumPage_AddBtn()
    # -------轉換視窗進行單據建檔--------------------
        handles = driver.window_handles
        # print(len(handles))
        driver.switch_to.window(handles[1])
        t0=driver.title
        print('目前窗口：'+t0)
        # 輸入單據金額
        test.input_PremiumPage_DocumentMoney(product)
        # 輸入使用金額
        test.input_PremiumPage_UseMoney(product)
        # 輸入立帳流水號
        test.input_PremiumPage_AccountNo(ASN)
        # 輸入繳款日期
        test.input_PremiumPage_EnteraccDate(personinfo)
        # 點擊保存按鈕
        test.click_PremiumPage_savebtn()
        # 選擇單據訊息結果的第一個選項
        test.click_PremiumPage_SelectResult()
        # 選擇繳費類型
        test.select_PremiumPage_PayType()
        # 選擇首期保費並全部提交
        test.select_PremiumPage_BusinessType(product,Policy_number)

    except Exception as e:
        print('有錯誤'+str(e))
        traceback.print_exc()
        # driver.close()
        # driver.quit()
    finally:
        print('測試結束')

if __name__ == "__main__":
    pytest.main()