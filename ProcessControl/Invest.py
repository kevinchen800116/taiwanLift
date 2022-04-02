from playwright.sync_api import Playwright
import os
import sys
o_path=os.path.abspath(os.path.join(os.getcwd(),'..'))
sys.path.append(o_path)

from PLClass import AddInfo

class Invest_Add(AddInfo):
    def __init__(self, playwright: Playwright):
        super().__init__(playwright)

            # 投資型商品錄入
    def Invest_productAdd(self, product):
        self.page2.touchscreen.tap(100,100)
        # 保存結束後要再按一次修改，不然投資標的的新增會失敗
        self.page2.frame(name="fraInterface").click("#updatebutton")
        self.page2.wait_for_timeout(5000)
        self.page2.frame(name="fraInterface").click("text=商品資訊錄入")
        self.page2.wait_for_timeout(2000)
        # 輸入險種編碼
        self.page2.frame(name="fraInterface").dblclick("input[name=\"RiskCode\"]")
        self.page2.frame(name="fraInterface").fill("input[name=\"RiskCode\"]", product["productId"])
        self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["productId"])
        self.page2.frame(name="fraInterface").press("input[name=\"RiskCode\"]", "Enter")
        # Click text=進入商品錄入畫面
        self.page2.frame(name="fraInterface").click("input[name=\"back\"]")
        self.page2.wait_for_timeout(2000)

        # (由此處開始修改)
        if(product["productId"]=="NIFA0802" or product["productId"]=="NIFA0801"):
            self.page2.frame(name="fraInterface").click("input[name=\"Prem\"]")
            self.page2.frame(name="fraInterface").fill("input[name=\"Prem\"]", product["price"])
            # 年金給付開始年齡
            self.page2.frame(name="fraInterface").click("input[name=\"GetYear\"]")
            self.page2.frame(name="fraInterface").fill("input[name=\"GetYear\"]", "45")
            # 年金給付方式
            self.page2.frame(name="fraInterface").dblclick("input[name=\"GetIntv\"]")

            self.page2.frame(name="fraInterface").fill("input[name=\"GetIntv\"]", "0")
            # 給付方式
            self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",value='0')

            self.page2.frame(name="fraInterface").press("input[name=\"GetIntv\"]", "Enter")

            # 幣別
            # CurrencyCode
            self.page2.frame(name="fraInterface").dblclick("input[name=\"CurrencyCode\"]")
            self.page2.frame(name="fraInterface").fill("input[name=\"CurrencyCode\"]", product["CurrencyCode"])
            # 選擇幣別
            self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["CurrencyCode"])
            self.page2.frame(name="fraInterface").press("input[name=\"CurrencyCode\"]", "Enter")

            # 保存
            self.page2.frame(name="fraInterface").click("text=保 存")
            # self.page2.wait_for_timeout(100000)
            
             # 選擇投標的，固定選E002
            if(product["investTarget1"] != product["investTarget2"]):
                self.page2.frame(name="fraInterface").dblclick("input[name=\"InvestPlanRate1\"]")
                self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["investTarget1"])
                self.page2.frame(name="fraInterface").press("input[name=\"InvestPlanRate1\"]", "Enter")
                # 輸入投資趴數
                self.page2.frame(name="fraInterface").fill("input[name=\"InvestPlanRate5\"]", product["InvestPlanRate"])
                # 輸入每月扣除順序
                self.page2.frame(name="fraInterface").fill("#InvestPlanRate7r0", "1")
                
                # 增加一欄
                self.page2.frame(name="fraInterface").click("input[name=\"InvestPlanRateaddOne\"]")
                # 第二欄投資標的
                self.page2.frame(name="fraInterface").dblclick("#InvestPlanRate1r1")
                self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["investTarget2"])
                self.page2.frame(name="fraInterface").press("#InvestPlanRate1r1", "Enter")
                # 投資趴數
                self.page2.frame(name="fraInterface").fill("#InvestPlanRate5r1", product["InvestPlanRate"])
                # 輸入每月扣除順序
                self.page2.frame(name="fraInterface").fill("#InvestPlanRate7r1", "2")

                # 添加
                self.page2.frame(name="fraInterface").click("input[name=\"submitFormButton\"]")
                # 上一步
                self.page2.frame(name="fraInterface").click("#riskbutton1")
            elif(product["investTarget1"] == product["investTarget2"]):
                self.page2.frame(name="fraInterface").dblclick("input[name=\"InvestPlanRate1\"]")
                self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["investTarget1"])
                self.page2.frame(name="fraInterface").press("input[name=\"InvestPlanRate1\"]", "Enter")
                # 輸入投資趴數
                self.page2.frame(name="fraInterface").fill("input[name=\"InvestPlanRate5\"]", "100")
                # 輸入每月扣除順序
                self.page2.frame(name="fraInterface").fill("#InvestPlanRate7r0", "1")
                # # 增加一欄
                # self.page2.frame(name="fraInterface").click("input[name=\"InvestPlanRateaddOne\"]")
                # # 第二欄投資標的
                # self.page2.frame(name="fraInterface").dblclick("#InvestPlanRate1r1")
                # self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["investTarget2"])
                # self.page2.frame(name="fraInterface").press("#InvestPlanRate1r1", "Enter")
                # # 投資趴數
                # self.page2.frame(name="fraInterface").fill("#InvestPlanRate5r1", product["InvestPlanRate"])
                # # 輸入每月扣除順序
                # self.page2.frame(name="fraInterface").fill("#InvestPlanRate7r1", "2")

                # 添加
                self.page2.frame(name="fraInterface").click("input[name=\"submitFormButton\"]")
                # 上一步
                self.page2.frame(name="fraInterface").click("#riskbutton1")

        # elif(product["productId"]!="NIFA0802"):
        else:
            # 輸入保額  Amnt
            self.page2.frame(name="fraInterface").click("#Amnt")
            self.page2.frame(name="fraInterface").fill("#Amnt", product["price"])

            # 選擇型別 PolType 甲乙型
            self.page2.frame(name="fraInterface").dblclick("input[name=\"PolType\"]")
            self.page2.frame(name="fraInterface").fill("input[name=\"PolType\"]", product["PolType"])
            self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["PolType"])
            self.page2.frame(name="fraInterface").press("input[name=\"PolType\"]", "Enter")

            ## 半年繳
            if (product["PayIntv"]=="6"):
                # 超額定期 class datagrid-editable-input
                self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[4]")
                # 此處element為選擇td的第三個框，若要選擇後方的繳別則需換成element[3]
                element=self.page2.frame(name="fraInterface").query_selector_all("input.datagrid-editable-input")
                element[2].fill(product["price2"])

                # 需跳出去雙擊其他欄位，解除這該死個輸入框鎖定後，再行對輸入框第一欄位輸入值
                self.page2.frame(name="fraInterface").dblclick("#Amnt")

                # 目標保費  class datagrid-editable-input
                self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[4]")
                element=self.page2.frame(name="fraInterface").query_selector_all("input.datagrid-editable-input")
                element[2].fill(product["price1"])
                # 繳別  combo-text validatebox-text
                self.page2.frame(name="fraInterface").dblclick("#Amnt")
                # 第一欄繳別
                self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[4]")
                # 點擊下拉選框
                self.page2.frame(name="fraInterface").click("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[5]/div/table/tbody/tr/td/span/span/span")
                # 強制點擊半年繳
                self.page2.frame(name="fraInterface").dispatch_event("//html/body/div[2]/div/div[3]","click")
                # 第二欄繳別
                self.page2.frame(name="fraInterface").dblclick("#Amnt")
                self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[5]")
                # 強制點擊半年繳
                self.page2.frame(name="fraInterface").dispatch_event("//html/body/div[3]/div/div[3]","click")
                # self.page2.frame(name="fraInterface").fill("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[5]/div/table/tbody/tr/td/span/input[1]",product["PayIntv"])
                # self.page2.frame(name="fraInterface").press("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[5]/div/table/tbody/tr/td/span/input[1]", "Enter")
                # 保存
                self.page2.frame(name="fraInterface").click("text=保 存")
                print("半年繳繳別選擇成功")
            ## 年繳
            elif (product["PayIntv"]=="12"):
                # 超額定期 class datagrid-editable-input
                # /html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[4]
                self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[4]")
                # 此處element為選擇td的第三個框，若要選擇後方的繳別則需換成element[3]
                element=self.page2.frame(name="fraInterface").query_selector_all("input.datagrid-editable-input")
                element[2].fill(product["price2"])

                # 需跳出去雙擊其他欄位，解除這該死個輸入框鎖定後，再行對輸入框第一欄位輸入值
                self.page2.frame(name="fraInterface").dblclick("#Amnt")

                # 目標保費  class datagrid-editable-input
                self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[4]")
                element=self.page2.frame(name="fraInterface").query_selector_all("input.datagrid-editable-input")
                element[2].fill(product["price1"])
                # 繳別  combo-text validatebox-text
                self.page2.frame(name="fraInterface").dblclick("#Amnt")
                # 第一欄繳別
                # /html/body/div[2]/div/div[5]
                self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[4]")
                # 點擊下拉選框
                self.page2.frame(name="fraInterface").click("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[5]/div/table/tbody/tr/td/span/span/span")
                # 強制點擊年繳
                self.page2.frame(name="fraInterface").dispatch_event("//html/body/div[2]/div/div[5]","click")
                # 第二欄繳別
                self.page2.frame(name="fraInterface").dblclick("#Amnt")
                self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[5]")
                # 強制點擊年繳
                self.page2.frame(name="fraInterface").dispatch_event("//html/body/div[3]/div/div[5]","click")
                # 保存
                self.page2.frame(name="fraInterface").click("text=保 存")
                print("年繳繳別選擇成功")
            ## 月繳
            elif (product["PayIntv"]=="1"):
                # 超額定期 class datagrid-editable-input
                # /html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[4]
                self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[4]")
                # 此處element為選擇td的第三個框，若要選擇後方的繳別則需換成element[3]
                element=self.page2.frame(name="fraInterface").query_selector_all("input.datagrid-editable-input")
                element[2].fill(product["price2"])

                # 需跳出去雙擊其他欄位，解除這該死個輸入框鎖定後，再行對輸入框第一欄位輸入值
                self.page2.frame(name="fraInterface").dblclick("#Amnt")

                # 目標保費  class datagrid-editable-input
                self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[4]")
                element=self.page2.frame(name="fraInterface").query_selector_all("input.datagrid-editable-input")
                element[2].fill(product["price1"])
                # 繳別  combo-text validatebox-text
                self.page2.frame(name="fraInterface").dblclick("#Amnt")
                # 第一欄繳別
                # /html/body/div[2]/div/div[5]
                self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[4]")
                # 點擊下拉選框
                self.page2.frame(name="fraInterface").click("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[5]/div/table/tbody/tr/td/span/span/span")
                # 強制點擊月繳
                self.page2.frame(name="fraInterface").dispatch_event("//html/body/div[2]/div/div[1]","click")
                # 第二欄繳別
                self.page2.frame(name="fraInterface").dblclick("#Amnt")
                self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[5]")
                # 強制點擊月繳
                self.page2.frame(name="fraInterface").dispatch_event("//html/body/div[3]/div/div[1]","click")
                # 保存
                self.page2.frame(name="fraInterface").click("text=保 存")
                print("月繳繳別選擇成功")
            ## 季繳
            elif (product["PayIntv"]=="3"):
                # 超額定期 class datagrid-editable-input
                # /html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[4]
                self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[4]")
                # 此處element為選擇td的第三個框，若要選擇後方的繳別則需換成element[3]
                element=self.page2.frame(name="fraInterface").query_selector_all("input.datagrid-editable-input")
                element[2].fill(product["price2"])

                # 需跳出去雙擊其他欄位，解除這該死個輸入框鎖定後，再行對輸入框第一欄位輸入值
                self.page2.frame(name="fraInterface").dblclick("#Amnt")

                # 目標保費  class datagrid-editable-input
                self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[4]")
                element=self.page2.frame(name="fraInterface").query_selector_all("input.datagrid-editable-input")
                element[2].fill(product["price1"])
                # 繳別  combo-text validatebox-text
                self.page2.frame(name="fraInterface").dblclick("#Amnt")
                # 第一欄繳別
                # /html/body/div[2]/div/div[5]
                self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[4]")
                # 點擊下拉選框
                self.page2.frame(name="fraInterface").click("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[5]/div/table/tbody/tr/td/span/span/span")
                # 強制點擊季繳
                self.page2.frame(name="fraInterface").dispatch_event("//html/body/div[2]/div/div[2]","click")
                # 第二欄繳別
                self.page2.frame(name="fraInterface").dblclick("#Amnt")
                self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[5]")
                # 強制點擊季繳
                self.page2.frame(name="fraInterface").dispatch_event("//html/body/div[3]/div/div[2]","click")
                # 保存
                self.page2.frame(name="fraInterface").click("text=保 存")
                print("季繳繳別選擇成功")
            #  躉繳
            elif(product["PayIntv"]=="0"):
                # 第一期保費責任(保費)
                self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[4]")
                # 此處element為選擇td的第三個框，若要選擇後方的繳別則需換成element[3]
                element=self.page2.frame(name="fraInterface").query_selector_all("input.datagrid-editable-input")
                element[2].fill(product["price1"])
                self.page2.frame(name="fraInterface").dblclick("#Amnt")
                # 第二欄繳別
                self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[5]")
                # 強制點擊躉繳
                self.page2.frame(name="fraInterface").dispatch_event("//html/body/div[2]/div/div[1]","click")
                # 保存
                self.page2.frame(name="fraInterface").click("text=保 存")
                # 需跳出去雙擊其他欄位，解除這該死個輸入框鎖定後，再行對輸入框第一欄位輸入值
                self.page2.frame(name="fraInterface").dblclick("#Amnt")
                print("躉繳繳別選擇成功")
            # 無繳別
            elif(product["PayIntv"]=="4"):
                # 第一期保費責任(保費)
                self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[4]")
                # 此處element為選擇td的第三個框，若要選擇後方的繳別則需換成element[3]
                element=self.page2.frame(name="fraInterface").query_selector_all("input.datagrid-editable-input")
                element[2].fill(product["price1"])
                # 需跳出去雙擊其他欄位，解除這該死個輸入框鎖定後，再行對輸入框第一欄位輸入值
                self.page2.frame(name="fraInterface").dblclick("#Amnt")
                self.page2.frame(name="fraInterface").click("text=保 存")
                print("無繳別選擇成功")
        
            print("繳別判斷已跳出")
            # 選擇銀行
            self.page2.frame(name="fraInterface").dblclick("#GetBankCode")
            self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",value="8220015")
            self.page2.frame(name="fraInterface").press("#GetBankCode", "Enter")
            # 輸入帳號
            self.page2.frame(name="fraInterface").fill("#GetBankAccNo", "123456789012")
            # 保 存
            self.page2.frame(name="fraInterface").click(":nth-match(:text('保 存'), 2)")

            # 選擇投標的，固定選E002
            if(product["investTarget1"] != product["investTarget2"]):
                self.page2.frame(name="fraInterface").dblclick("input[name=\"InvestPlanRate1\"]")
                self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["investTarget1"])
                self.page2.frame(name="fraInterface").press("input[name=\"InvestPlanRate1\"]", "Enter")
                # 輸入投資趴數
                self.page2.frame(name="fraInterface").fill("input[name=\"InvestPlanRate5\"]", product["InvestPlanRate"])
                # 輸入每月扣除順序
                self.page2.frame(name="fraInterface").fill("#InvestPlanRate7r0", "1")
                
                # 增加一欄
                self.page2.frame(name="fraInterface").click("input[name=\"InvestPlanRateaddOne\"]")
                # 第二欄投資標的
                self.page2.frame(name="fraInterface").dblclick("#InvestPlanRate1r1")
                self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["investTarget2"])
                self.page2.frame(name="fraInterface").press("#InvestPlanRate1r1", "Enter")
                # 投資趴數
                self.page2.frame(name="fraInterface").fill("#InvestPlanRate5r1", product["InvestPlanRate"])
                # 輸入每月扣除順序
                self.page2.frame(name="fraInterface").fill("#InvestPlanRate7r1", "2")

                # 添加
                self.page2.frame(name="fraInterface").click("input[name=\"submitFormButton\"]")
                # 上一步
                self.page2.frame(name="fraInterface").click("#riskbutton1")
            elif(product["investTarget1"] == product["investTarget2"]):
                self.page2.frame(name="fraInterface").dblclick("input[name=\"InvestPlanRate1\"]")
                self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["investTarget1"])
                self.page2.frame(name="fraInterface").press("input[name=\"InvestPlanRate1\"]", "Enter")
                # 輸入投資趴數
                self.page2.frame(name="fraInterface").fill("input[name=\"InvestPlanRate5\"]", "100")
                # 輸入每月扣除順序
                self.page2.frame(name="fraInterface").fill("#InvestPlanRate7r0", "1")
                # # 增加一欄
                # self.page2.frame(name="fraInterface").click("input[name=\"InvestPlanRateaddOne\"]")
                # # 第二欄投資標的
                # self.page2.frame(name="fraInterface").dblclick("#InvestPlanRate1r1")
                # self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["investTarget2"])
                # self.page2.frame(name="fraInterface").press("#InvestPlanRate1r1", "Enter")
                # # 投資趴數
                # self.page2.frame(name="fraInterface").fill("#InvestPlanRate5r1", product["InvestPlanRate"])
                # # 輸入每月扣除順序
                # self.page2.frame(name="fraInterface").fill("#InvestPlanRate7r1", "2")

                # 添加
                self.page2.frame(name="fraInterface").click("input[name=\"submitFormButton\"]")
                # 上一步
                self.page2.frame(name="fraInterface").click("#riskbutton1")

    def Invest_BenefitPeople(self, BenefitInfo):
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

    def Invest_ImportInfo(self):
        self.page2.touchscreen.tap(100,100)
        # self.page2.screenshot(path="重要事項告知書.png")
        self.page2.once("dialog", lambda dialog: dialog.accept())
        with self.page2.expect_popup() as popup_info:
            self.page2.frame(name="fraInterface").click("#ImpInf")
        self.page7 = popup_info.value
        self.page7.set_viewport_size({"width": 1920, "height": 1080})
        # self.page7.screenshot(path="重要事項告知書2.png")
        self.page7.touchscreen.tap(100,100)
        print(self.page7.title())

        # (全部打勾)
        # self.page7.frame(name="fraInterface").dispatch_event("//html/body/form/div[3]/div/table/tbody/tr/td/span/div[1]/div[3]/div/table/thead/tr/th[1]/div/input","click")
        self.page7.frame(name="fraInterface").check("#ImportInformGridChk0")
        self.page7.frame(name="fraInterface").check("#ImportInformGridChk1")
        self.page7.frame(name="fraInterface").check("#ImportInformGridChk2")
        self.page7.frame(name="fraInterface").click("#save")
        self.page7.once("dialog", lambda dialog: dialog.dismiss())
        self.page7.wait_for_timeout(2000)

        self.page7.frame(name="fraInterface").click("text=返回")