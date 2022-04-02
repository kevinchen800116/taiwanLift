from playwright.sync_api import sync_playwright
from selenium import webdriver
import pytest
import traceback
import os
import sys

### 移動到ProcessControl的資料夾(此動作是為了取得Tranditional的py檔案內的Tranditional_Add()的class)
o_path=os.path.abspath(os.path.join(os.getcwd(),'../ProcessControl'))
sys.path.append(o_path)

# from Invest import Invest_Add
from Invest import Invest_Add
# from NITU0905_class import NITU0905_Add


# 執行此命令進行測試
# python -m pytest -v -s

def test_NIFA0801(my_fixture):
    ### 依照All_test_Data的append順序決定
    personinfo=my_fixture[0]
    BenefitInfo=my_fixture[3]
    SaleReport=my_fixture[1]
    product=my_fixture[2]
    # print(product["price"])

    with sync_playwright() as playwright:
        try:
            test= Invest_Add(playwright)
            # test = AddInfo(playwright)
            test.navigate()
            test.login('TEST41', 'admin001')
            test.SetTime(personinfo)

    # ## 新增
            # test.Noscan_apply()
            # # # 要保人資訊新增
            # test.Add(personinfo)
    ##查詢
            test.NoScan_query(personinfo)
            # 要保人資訊新增
            Salesman_value=test.Add(personinfo)

    # # # 投資商品新增(傳統型)

            # # # 投資型商品資訊錄入
            test.Invest_productAdd(product)

            # # 業務員招攬報告書
            test.SalesmanReport(SaleReport)

            # # 銀行轉帳授權書(等於0代表手續期都是自行繳費)
            if(personinfo["AuthorType"] != "0"):
                test.Authorization_Letter(personinfo, Salesman_value)

            # # 共同行銷及特定目的外搜集、處理、利用個人資料之聲明同意書
            test.commonpurpose()
            # 投資屬性分析問券
            test.InvestReport(personinfo)

            # # # 投資型受益人
            test.Invest_BenefitPeople(BenefitInfo)
            # 重要事項告知書
            test.Invest_ImportInfo()
            test.FATCA(personinfo)
            test.CRS(personinfo)
            # test.tranditional_contract_other_info(product)


            test.logout()


        except Exception as e:
            # 在終端機中打印錯誤信息，以紅色顯示。
            traceback.print_exc()
            print('playwright有錯誤:'+ str(e))
            # 使用變量接錯誤訊息
            # err=traceback.format_exc()
            test.logout()
        finally:
            print('playwright測試結束了')

if __name__ == "__main__":
    pytest.main()
