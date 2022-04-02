from playwright.sync_api import Playwright
import os
import sys
o_path=os.path.abspath(os.path.join(os.getcwd(),'..'))
sys.path.append(o_path)

from PLClass import AddInfo

class Tranditional_Add(AddInfo):
    def __init__(self, playwright: Playwright):
        super().__init__(playwright)

        # 傳統型商品錄入 
    def tranditional_productAdd(self,product):
            self.page2.touchscreen.tap(100,100)
            # 保存結束後要再按一次修改，不然投資標的的新增會失敗
            # self.page2.frame(name="fraInterface").click("#updatebutton")
            # self.page2.wait_for_timeout(5000)
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

            # 增值回饋分享金給付方式
            if(product["productId"] != 'NTWL0105'):
                self.page2.frame(name="fraInterface").dblclick("input[name=\"BonusGetMode\"]")
                self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]","1")
                self.page2.frame(name="fraInterface").press("input[name=\"BonusGetMode\"]", "Enter")


            # 自動墊繳
            self.page2.frame(name="fraInterface").dblclick("input[name=\"AutoPayFlag\"]")
            self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["AutoPayFlag"])
            self.page2.frame(name="fraInterface").press("input[name=\"AutoPayFlag\"]", "Enter")

            # 保存
            self.page2.frame(name="fraInterface").click("text=保 存")
            self.page2.wait_for_timeout(2000)

            # 修改
            # self.page2.frame(name="fraInterface").click("text=修  改")

            # 上一步
            self.page2.frame(name="fraInterface").click("#riskbutton1")
        

    def Tranditional_BenefitPeople(self, BenefitInfo):
        # time.sleep(2)
        # print("進入自己的投資方法")
        self.page2.screenshot(path="受益人資料錄入2.png")
        self.page2.touchscreen.tap(100,100)
        # self.page2.screenshot(path="受益人資料錄入.png")
        self.page2.once("dialog", lambda dialog: dialog.accept())
        with self.page2.expect_popup() as popup_info:
            self.page2.frame(name="fraInterface").click("#LBnfButton")
        self.page6 = popup_info.value
        self.page6.set_viewport_size({"width": 1920, "height": 1080})

        
        # # 受益人類別
        # BnfType 身故受益人
        self.page6.touchscreen.tap(100,100)
        self.page6.frame(name="fraInterface").fill("#BnfType", BenefitInfo["BnfType1"])
        self.page6.frame(name="fraInterface").dblclick("#BnfType")
        self.page6.wait_for_timeout(3000)
        self.page6.frame(name="fraInterface").click("option[value=\"1\"]")
        self.page6.wait_for_timeout(2000)

        # # 法定 check
        # isSameLaw
        self.page6.frame(name="fraInterface").check("#isSameLaw")
        # # 受益人順位
        # BnfGrade 第一順位
        self.page6.frame(name="fraInterface").dblclick("#BnfGrade")
        self.page6.frame(name="fraInterface").fill("#BnfGrade", BenefitInfo["BnfGrade"])
        self.page6.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",BenefitInfo["BnfGrade"])
        self.page6.wait_for_timeout(3000)

        self.page6.frame(name="fraInterface").press("#BnfGrade", "Enter")
        # 100%
        self.page6.frame(name="fraInterface").fill("#BnfLot", BenefitInfo["BnfLot"])       

        # 保存
        self.page6.frame(name="fraInterface").click("#save")
        self.page6.once("dialog", lambda dialog: dialog.dismiss())
        self.page6.frame(name="fraInterface").click("text=返回")

    # 分期定期給付
    def InstallPayButton(self):
        self.page2.touchscreen.tap(100,100)
        self.page2.once("dialog", lambda dialog: dialog.accept())
        with self.page2.expect_popup() as popup_info:
            self.page2.frame(name="fraInterface").click("#InstallPayButton")
        self.page18 = popup_info.value
        self.page18.wait_for_load_state()
        print(self.page18.title())
        
        # 身故/完全失能/特定意外傷害第一級失能保險金執行給付方式
        self.page18.frame(name="fraInterface").click("#PayMode")
        self.page18.frame(name="fraInterface").fill("#PayMode", "0")
        self.page18.frame(name="fraInterface").select_option("select[name=\"codeselect\"]", "0")
        self.page18.frame(name="fraInterface").press("#PayMode", "Enter")

        # 新增
        self.page18.frame(name="fraInterface").click("#saveInstallmentPayBtn") 

        # 返回
        self.page18.frame(name="fraInterface").click("#riskbutton1")



        # if(product["productId"]=="NIFA0802"):
        #     self.page2.frame(name="fraInterface").click("text=商品資訊錄入")
        #     self.page2.wait_for_timeout(2000)

        #     self.page2.frame(name="fraInterface").dblclick("input[name=\"RiskCode\"]")
        #     self.page2.frame(name="fraInterface").fill("input[name=\"RiskCode\"]", product["productId"])

        #     self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["productId"])

        #     self.page2.frame(name="fraInterface").press("input[name=\"RiskCode\"]", "Enter")

        #     # Click text=進入商品錄入畫面
        #     self.page2.frame(name="fraInterface").click("input[name=\"back\"]")
        #     self.page2.wait_for_timeout(2000)

        #     self.page2.frame(name="fraInterface").click("input[name=\"Prem\"]")
        #     self.page2.frame(name="fraInterface").fill("input[name=\"Prem\"]", product["price"])
        #     # 年金給付開始年齡
        #     self.page2.frame(name="fraInterface").click("input[name=\"GetYear\"]")
        #     self.page2.frame(name="fraInterface").fill("input[name=\"GetYear\"]", "45")
        #     # 年金給付方式
        #     self.page2.frame(name="fraInterface").dblclick("input[name=\"GetIntv\"]")

        #     self.page2.frame(name="fraInterface").fill("input[name=\"GetIntv\"]", "0")
        #     # 給付方式
        #     self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",value='0')

        #     self.page2.frame(name="fraInterface").press("input[name=\"GetIntv\"]", "Enter")

        #     # 幣別
        #     # CurrencyCode
        #     self.page2.frame(name="fraInterface").dblclick("input[name=\"CurrencyCode\"]")
        #     self.page2.frame(name="fraInterface").fill("input[name=\"CurrencyCode\"]", product["CurrencyCode"])
        #     # 選擇幣別
        #     self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["CurrencyCode"])
        #     self.page2.frame(name="fraInterface").press("input[name=\"CurrencyCode\"]", "Enter")

        #     # 保存
        #     self.page2.frame(name="fraInterface").click("text=保 存")
        #     # self.page2.wait_for_timeout(100000)
            
        #     # 上一步
        #     self.page2.frame(name="fraInterface").click("#riskbutton1")

        # elif(product["productId"]!="NIFA0802"):