from selenium import webdriver
from U_LoginPage import LoginPage
import traceback
import pytest
import os
import sys
import time
import datetime
o_path = os.path.abspath(os.path.join(os.getcwd(), '..'))
sys.path.append(o_path)


# 執行此命令進行測試
# python -m pytest -v -s

def test_NIFA0801(my_fixture):
    # 依照All_test_Data的append順序決定
    personinfo = my_fixture[0]
    BenefitInfo = my_fixture[3]
    SaleReport = my_fixture[1]
    product = my_fixture[2]
    # print(product["price"])
    try:
        options = webdriver.ChromeOptions()
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        options.add_argument('--ignore-certificate-errors')
        options.add_experimental_option(
            "excludeSwitches", ['enable-automation'])
        driver = webdriver.Chrome(options=options)
        driver.implicitly_wait(30)

        # url = "https://10.1.113.23:9443/"
        url = "https://www.google.com/"
        username = "TEST36"
        password = "admin001"

        # Policy_number='9005009993'# (女)9005010510

    # ------登入---------
        test = LoginPage(driver, url, u"人壽保險核心業務系統")
        test.open()
        test.switch_frame('fraInterface')
        test.input_username(username)
        test.input_password(password)
        test.click_submit()
        print('登入結束')

        test.mouse_To_SystemTime(personinfo)

        test.mouse_To_NewIInput('88ANB11101119001d043')

        time.sleep(1)
        test.switch_targetWindow("扫描件显示")

        # test.NewInput_personinfo_input(personinfo)

        test.NewInput_personinfo_input2(personinfo)

# ----------------------------------------------------------
        # 無特定險種代碼
        # test.NewInput_Product_input(product)

        # 台灣人壽愛豐收外幣變額萬能壽險
        # test.product_NIFU0501(product)

        # 台灣人壽金采100變額年金保險
        # test.product_NITA0901(product)

        # 2308 台灣人壽金采100外幣變額萬能壽險
        # test.product_NIFU0101(product)

        # 2315台灣人壽月月好鑫外幣變額年金保險要保書
        # test.product_NIFA1602(product)

        # 2316台灣人壽月月好鑫外幣變額年金保險
        # test.product_NIFU1203(product)

        # All
        # test.product(product)

# ----------------------------------------------------------
        # test.NewInput_otherInfo_input()

        # test.NewInput_paymentInfo_input()

        # test.BenefitPeople()

        # test.InstallPay()

    except Exception as e:
        print('有錯誤'+str(e))
        traceback.print_exc()
        # driver.quit()

    finally:
        print('測試結束')
        # driver.quit()


if __name__ == "__main__":
    pytest.main()
