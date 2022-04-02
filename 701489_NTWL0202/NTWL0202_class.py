from playwright.sync_api import Playwright
import os
import sys
o_path=os.path.abspath(os.path.join(os.getcwd(),'..'))
sys.path.append(o_path)

from PLClass import AddInfo


class NTWL0202_Add(AddInfo):
    def __init__(self, playwright: Playwright):
        super().__init__(playwright)
        
    def tranditional_productAdd_1113_NTWL0202(self,product):
        self.page2.touchscreen.tap(100,100)
        # 保存結束後要再按一次修改，不然投資標的的新增會失敗
        self.page2.frame(name="fraInterface").click("#updatebutton")
        self.page2.wait_for_timeout(5000)
        self.page2.frame(name="fraInterface").click("text=商品資訊錄入")
        self.page2.wait_for_timeout(2000)

        self.page2.frame(name="fraInterface").dblclick("input[name=\"RiskCode\"]")
        self.page2.frame(name="fraInterface").fill("input[name=\"RiskCode\"]", product["productId"])

        self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["productId"])

        self.page2.frame(name="fraInterface").press("input[name=\"RiskCode\"]", "Enter")

        # Click text=進入商品錄入畫面
        self.page2.frame(name="fraInterface").click("input[name=\"back\"]")
        self.page2.wait_for_timeout(2000)

        # # 保費(會自動算出不需要keyin)
        # self.page2.frame(name="fraInterface").click("input[name=\"Prem\"]")
        # self.page2.frame(name="fraInterface").fill("input[name=\"Prem\"]", product["price"])

        # 保額
        self.page2.frame(name="fraInterface").click("input[name=\"Amnt\"]")
        self.page2.frame(name="fraInterface").fill("input[name=\"Amnt\"]", product["price"])

        # 繳別
        self.page2.frame(name="fraInterface").dblclick("input[name=\"PayIntv\"]")
        self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["PayIntv"])
        self.page2.frame(name="fraInterface").press("input[name=\"PayIntv\"]", "Enter")

        # 繳費期間
        self.page2.frame(name="fraInterface").dblclick("input[name=\"PayEndYear\"]")
        self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["PayEndYear"])
        self.page2.frame(name="fraInterface").press("input[name=\"PayEndYear\"]", "Enter")


        # 自動墊繳
        self.page2.frame(name="fraInterface").dblclick("input[name=\"AutoPayFlag\"]")
        self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["AutoPayFlag"])
        self.page2.frame(name="fraInterface").press("input[name=\"AutoPayFlag\"]", "Enter")

        # 保存
        self.page2.frame(name="fraInterface").click("text=保 存")
        self.page2.wait_for_timeout(2000)

        # # 獲取保費
        # handle = self.page2.query_selector("input[name=\"Prem\"]")
        # value = handle.get_attribute("value")
        # 上一步
        self.page2.frame(name="fraInterface").click("#riskbutton1")

