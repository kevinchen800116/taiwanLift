
from playwright.sync_api import Playwright
import os
import sys
o_path=os.path.abspath(os.path.join(os.getcwd(),'..'))
sys.path.append(o_path)

from PLClass import AddInfo


class NITU0905_Add(AddInfo): # 從AddInfo子繼承過來 
    def __init__(self, playwright: Playwright):
        super().__init__(playwright)

        # 投資型商品錄入
    # def NITU0905_Add(self, product):
    #     self.page2.touchscreen.tap(100,100)
    #     # 保存結束後要再按一次修改，不然投資標的的新增會失敗
    #     self.page2.frame(name="fraInterface").click("#updatebutton")
    #     self.page2.wait_for_timeout(5000)
    #     self.page2.frame(name="fraInterface").click("text=商品資訊錄入")
    #     self.page2.wait_for_timeout(2000)
    #     # 輸入險種編碼
    #     self.page2.frame(name="fraInterface").dblclick("input[name=\"RiskCode\"]")
    #     self.page2.frame(name="fraInterface").fill("input[name=\"RiskCode\"]", product["productId"])
    #     self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["productId"])
    #     self.page2.frame(name="fraInterface").press("input[name=\"RiskCode\"]", "Enter")
    #     # Click text=進入商品錄入畫面
    #     self.page2.frame(name="fraInterface").click("input[name=\"back\"]")
    #     self.page2.wait_for_timeout(2000)

    #     # (由此處開始修改)
    #     # 輸入保額  Amnt
    #     self.page2.frame(name="fraInterface").click("#Amnt")
    #     self.page2.frame(name="fraInterface").fill("#Amnt", product["price"])

    #     # 選擇型別 PolType 甲乙型
    #     self.page2.frame(name="fraInterface").dblclick("input[name=\"PolType\"]")
    #     self.page2.frame(name="fraInterface").fill("input[name=\"PolType\"]", product["PolType"])
    #     self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["PolType"])
    #     self.page2.frame(name="fraInterface").press("input[name=\"PolType\"]", "Enter")


    #     # 超額定期 class datagrid-editable-input
    #     # /html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[4]
    #     self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[4]")
    #     # 此處element為選擇td的第三個框，若要選擇後方的繳別則需換成element[3]
    #     element=self.page2.frame(name="fraInterface").query_selector_all("input.datagrid-editable-input")
    #     element[2].fill(product["price2"])

    #     # 需跳出去雙擊其他欄位，解除這該死個輸入框鎖定後，再行對輸入框第一欄位輸入值
    #     self.page2.frame(name="fraInterface").dblclick("#Amnt")

    #     # 目標保費  class datagrid-editable-input
    #     self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[4]")
    #     element=self.page2.frame(name="fraInterface").query_selector_all("input.datagrid-editable-input")
    #     element[2].fill(product["price1"])

    #     if (product["PayIntv"]=="6"):
    #         # 繳別  combo-text validatebox-text
    #         self.page2.frame(name="fraInterface").dblclick("#Amnt")
    #         # 第一欄繳別
    #         self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[4]")
    #         # 點擊下拉選框
    #         self.page2.frame(name="fraInterface").click("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[5]/div/table/tbody/tr/td/span/span/span")
    #         # 強制點擊半年繳
    #         self.page2.frame(name="fraInterface").dispatch_event("//html/body/div[2]/div/div[3]","click")
    #         # 第二欄繳別
    #         self.page2.frame(name="fraInterface").dblclick("#Amnt")
    #         self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[5]")
    #         # 強制點擊半年繳
    #         self.page2.frame(name="fraInterface").dispatch_event("//html/body/div[3]/div/div[3]","click")
    #         # self.page2.frame(name="fraInterface").fill("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[5]/div/table/tbody/tr/td/span/input[1]",product["PayIntv"])
    #         # self.page2.frame(name="fraInterface").press("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[5]/div/table/tbody/tr/td/span/input[1]", "Enter")
    #         # 保存
    #         self.page2.frame(name="fraInterface").click("text=保 存")
    #     elif (product["PayIntv"]=="12"):
    #         # 繳別  combo-text validatebox-text
    #         self.page2.frame(name="fraInterface").dblclick("#Amnt")
    #         # 第一欄繳別
    #         # /html/body/div[2]/div/div[5]
    #         self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[4]")
    #         # 點擊下拉選框
    #         self.page2.frame(name="fraInterface").click("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[5]/div/table/tbody/tr/td/span/span/span")
    #         # 強制點擊年繳
    #         self.page2.frame(name="fraInterface").dispatch_event("//html/body/div[2]/div/div[5]","click")
    #         # 第二欄繳別
    #         self.page2.frame(name="fraInterface").dblclick("#Amnt")
    #         self.page2.frame(name="fraInterface").dblclick("//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[2]/td[5]")
    #         # 強制點擊年繳
    #         self.page2.frame(name="fraInterface").dispatch_event("//html/body/div[3]/div/div[5]","click")
    #         # 保存
    #         self.page2.frame(name="fraInterface").click("text=保 存")
        
    #     print("繳別選擇成功")
    #     # 選擇銀行
    #     self.page2.frame(name="fraInterface").dblclick("#GetBankCode")
    #     self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",value="8220015")
    #     self.page2.frame(name="fraInterface").press("#GetBankCode", "Enter")
    #     # 輸入帳號
    #     self.page2.frame(name="fraInterface").fill("#GetBankAccNo", "123456789012")
    #     # 保 存
    #     self.page2.frame(name="fraInterface").click(":nth-match(:text('保 存'), 2)")

    #     # 選擇投標的，固定選E002
    #     self.page2.frame(name="fraInterface").dblclick("input[name=\"InvestPlanRate1\"]")
    #     self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",value="M001")
    #     self.page2.frame(name="fraInterface").press("input[name=\"InvestPlanRate1\"]", "Enter")
    #     # 輸入投資趴數
    #     self.page2.frame(name="fraInterface").fill("input[name=\"InvestPlanRate5\"]", product["InvestPlanRate"])
    #     # 輸入每月扣除順序
    #     self.page2.frame(name="fraInterface").fill("#InvestPlanRate7r0", "1")

    #     # 增加一欄
    #     self.page2.frame(name="fraInterface").click("input[name=\"InvestPlanRateaddOne\"]")
    #     # 第二欄投資標的
    #     self.page2.frame(name="fraInterface").dblclick("#InvestPlanRate1r1")
    #     self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",value="ABHBM")
    #     self.page2.frame(name="fraInterface").press("#InvestPlanRate1r1", "Enter")
    #     # 第二攔投資趴數
    #     self.page2.frame(name="fraInterface").fill("#InvestPlanRate5r1", product["InvestPlanRate"])
    #     # 輸入每月扣除順序
    #     self.page2.frame(name="fraInterface").fill("#InvestPlanRate7r1", "2")


    #     # 添加
    #     self.page2.frame(name="fraInterface").click("input[name=\"submitFormButton\"]")
    #     # 上一步
    #     self.page2.frame(name="fraInterface").click("#riskbutton1")

    # def BenefitPeople_NITU0905_NTD(self, BenefitInfo):
    #     # time.sleep(2)
    #     # print("進入自己的投資方法")
    #     self.page2.screenshot(path="受益人資料錄入2.png")
    #     self.page2.touchscreen.tap(100,100)
    #     # self.page2.screenshot(path="受益人資料錄入.png")
    #     self.page2.once("dialog", lambda dialog: dialog.accept())
    #     with self.page2.expect_popup() as popup_info:
    #         self.page2.frame(name="fraInterface").click("#LBnfButton")
    #     self.page6 = popup_info.value
    #     self.page6.set_viewport_size({"width": 1920, "height": 1080})

        
    #     # # 受益人類別
    #     # BnfType 身故受益人
    #     self.page6.touchscreen.tap(100,100)
    #     self.page6.frame(name="fraInterface").fill("#BnfType", BenefitInfo["BnfType1"])
    #     self.page6.frame(name="fraInterface").dblclick("#BnfType")
    #     self.page6.wait_for_timeout(3000)
    #     self.page6.frame(name="fraInterface").click("option[value=\"1\"]")
    #     self.page6.wait_for_timeout(2000)

    #     # # 法定 check
    #     # isSameLaw
    #     self.page6.frame(name="fraInterface").check("#isSameLaw")
    #     # # 受益人順位
    #     # BnfGrade 第一順位
    #     self.page6.frame(name="fraInterface").dblclick("#BnfGrade")
    #     self.page6.frame(name="fraInterface").fill("#BnfGrade", BenefitInfo["BnfGrade"])
    #     self.page6.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",BenefitInfo["BnfGrade"])
    #     self.page6.wait_for_timeout(3000)

    #     self.page6.frame(name="fraInterface").press("#BnfGrade", "Enter")
    #     # 100%
    #     self.page6.frame(name="fraInterface").fill("#BnfLot", BenefitInfo["BnfLot"])       

    #     # 保存
    #     self.page6.frame(name="fraInterface").click("#save")
    #     self.page6.once("dialog", lambda dialog: dialog.dismiss())
    #     self.page6.frame(name="fraInterface").click("text=返回")

    # def ImportInf_NITU0905_NTD(self):
    #     print("NITU0901_NTD的重要事項告知書")
    #     self.page2.touchscreen.tap(100,100)
    #     # self.page2.screenshot(path="重要事項告知書.png")
    #     self.page2.once("dialog", lambda dialog: dialog.accept())
    #     with self.page2.expect_popup() as popup_info:
    #         self.page2.frame(name="fraInterface").click("#ImpInf")
    #     self.page7 = popup_info.value
    #     self.page7.set_viewport_size({"width": 1920, "height": 1080})
    #     # self.page7.screenshot(path="重要事項告知書2.png")
    #     self.page7.touchscreen.tap(100,100)
    #     print(self.page7.title())

    #     # (全部打勾)
    #     # self.page7.frame(name="fraInterface").dispatch_event("//html/body/form/div[3]/div/table/tbody/tr/td/span/div[1]/div[3]/div/table/thead/tr/th[1]/div/input","click")
    #     self.page7.frame(name="fraInterface").check("#ImportInformGridChk0")
    #     self.page7.frame(name="fraInterface").check("#ImportInformGridChk1")
    #     self.page7.frame(name="fraInterface").check("#ImportInformGridChk2")
    #     self.page7.frame(name="fraInterface").click("#save")
    #     self.page7.once("dialog", lambda dialog: dialog.dismiss())
    #     self.page7.wait_for_timeout(2000)

    #     self.page7.frame(name="fraInterface").click("text=返回")

    # def InvestReport_NITU0905(self,personinfo):
    #     self.page2.touchscreen.tap(100,100)
    #     self.page2.once("dialog", lambda dialog: dialog.accept())
    #     with self.page2.expect_popup() as popup_info:
    #         self.page2.frame(name="fraInterface").click("#InAttQue")
    #     self.page8 = popup_info.value
    #     self.page8.set_viewport_size({"width": 1920, "height": 1080})
    #     # self.page8.touchscreen.tap(100,100)
    #     self.page8.screenshot(path="投資屬性分析問卷.png")
    #     print(self.page8.title())
        
    #     if(20<int(personinfo['age'])<=44):
    #         # 20-44
    #         print("20歲以上")
    #         self.page8.frame(name="fraInterface").check("#AgeGridSel1")
    #         self.page8.frame(name="fraInterface").check("#JobGridSel0")
    #     elif(45<int(personinfo['age'])<=65):
    #         # 45-65
    #         print("45歲以上")
    #         self.page8.frame(name="fraInterface").check("#AgeGridSel2")
    #         self.page8.frame(name="fraInterface").check("#JobGridSel0")
    #     elif(66<int(personinfo['age'])<=69):
    #         # 66-69
    #         print("66歲以上")
    #         self.page8.frame(name="fraInterface").check("#AgeGridSel3")
    #         self.page8.frame(name="fraInterface").check("#JobGridSel0")
    #     elif(int(personinfo['age'])>70):
    #         # >70
    #         print("70歲以上")
    #         self.page8.frame(name="fraInterface").check("#AgeGridSel4")
    #         self.page8.frame(name="fraInterface").check("#JobGridSel0")

    #     self.page8.frame(name="fraInterface").check("#ExeGridChk2")
    #     self.page8.frame(name="fraInterface").check("#ExeGridChk4")
    #     self.page8.frame(name="fraInterface").check("#ExeGridChk6")

    #     self.page8.frame(name="fraInterface").check("#RiskGridSel0")

    #     if(20<int(personinfo['age'])<=44):
    #         self.page8.frame(name="fraInterface").fill("#Total",value='45')
    #     elif(45<int(personinfo['age'])<=65):
    #         self.page8.frame(name="fraInterface").fill("#Total",value='47')
    #     elif(66<int(personinfo['age'])<=69):
    #         self.page8.frame(name="fraInterface").fill("#Total",value='43')
    #     elif(int(personinfo['age'])>70):
    #         self.page8.frame(name="fraInterface").fill("#Total",value='42')

    #     # 同要保書電話
    #     self.page8.frame(name="fraInterface").dblclick("#CallPhoneFlag")
    #     self.page8.frame(name="fraInterface").fill("#CallPhoneFlag", "1")
    #     self.page8.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",value='1')
    #     self.page8.wait_for_timeout(3000)
    #     self.page8.frame(name="fraInterface").press("#CallPhoneFlag", "Enter")
    #     # 投資人須知]、[保險商品說明書]交付確認書
    #     self.page8.frame(name="fraInterface").check("#Confirm1GridChk0")
    #     self.page8.frame(name="fraInterface").check("#Confirm1GridChk1")
    #     self.page8.frame(name="fraInterface").check("#Confirm1GridChk2")
    #     self.page8.frame(name="fraInterface").check("#Confirm1GridChk3")

    #     self.page8.frame(name="fraInterface").check("#Confirm2GridChk0")
    #     self.page8.frame(name="fraInterface").check("#Confirm2GridChk1")
        
    #     self.page8.frame(name="fraInterface").click("#save")
    #     self.page8.wait_for_timeout(2000)
    #     # self.page8.screenshot(path="投資屬性分析問卷2.png")
    #     self.page8.frame(name="fraInterface").click("text=返回")

