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
        Policy_number=""

        
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

        test.mouse_To_BankTransfer_Return()
        if(personinfo["bankcode"]=="1080014"):
          # 5.使用send_keys方法上傳檔案
          test.upload_BF_file(personinfo)
          test.click_BF_saveBtn(personinfo)
          test.take_screenshot()
        else:
          test.select_BR_result()
          test.input_BR_AuthorBar(Policy_number)
          test.click_BF_submitBtn()
          test.take_screenshot()
          time.sleep(1)

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