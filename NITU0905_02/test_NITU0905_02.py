from playwright.sync_api import sync_playwright
from selenium import webdriver
import pytest
import traceback
# import os
# import sys
# # o_path = os.getcwd() # 返回當前工作目錄
# # print ('我是'+os.path.abspath(os.path.join(os.getcwd(),'..')))
# o_path=os.path.abspath(os.path.join(os.getcwd(),'..'))
# sys.path.append(o_path)

from NITU0905_02_class import NITU0905_02_Add


# 執行此命令進行測試
# python -m pytest -v -s

def test_NITU0905_02(my_fixture):
    personinfo=my_fixture[0]
    BenefitInfo=my_fixture[1]
    SaleReport=my_fixture[2]
    product=my_fixture[3]
    # print(product["price"])

    with sync_playwright() as playwright:
        try:
            test= NITU0905_02_Add(playwright)
            # test = AddInfo(playwright)
            test.navigate()
            test.login('TEST41', 'admin001')

    # ## 新增
            # test.Noscan_apply()
            # # # 要保人資訊新增
            # test.Add(personinfo)

    ##查詢
            test.NoScan_query(personinfo)
            # # 要保人資訊新增
            test.Add(personinfo)

    # # # 投資商品新增(傳統型)

            # # # 傳統型商品資訊錄入
            test.NITU0905_02_Add(product)

            # # # 業務員招攬報告書
            test.SalesmanReport(SaleReport)

            # # # 共同行銷及特定目的外搜集、處理、利用個人資料之聲明同意書
            test.commonpurpose()

            # # # 傳統型受益人
            test.BenefitPeople_NITU0905_02_NTD(BenefitInfo)
            test.FATCA(personinfo)
            test.ImportInf_NITU0905_02_NTD()
            test.InvestReport_NITU0905_02(personinfo)
            test.Risk(SaleReport)
            test.CRS(personinfo)
            test.contract_other_info()


            test.logout()


        except Exception as e:
            print('playwright有錯誤:'+ str(e))
            # 在終端機中打印錯誤信息，以紅色顯示。
            traceback.print_exc()
            # 使用變量接錯誤訊息
            # err=traceback.format_exc()
            test.logout()
        finally:
            print('playwright測試結束了')

if __name__ == "__main__":
    pytest.main()
