from playwright.sync_api import Playwright, sync_playwright
import time
# from U_LoginPage import LoginPage


class LoginPage():
    # 以playwright的變數形式將Playwright的類別載入進來
    def __init__(self, playwright: Playwright):
        # 初始化playwright物件的設定
        browser = playwright.chromium.launch(headless=False)
        context = browser.new_context(ignore_https_errors=True, has_touch=True, viewport={
                                      "width": 1920, "height": 1080})
        page = context.new_page()
        self.browser = browser
        self.context = context
        self.page = page

    def navigate(self):
        self.page.goto("https://10.1.113.23:9443/")

        # playwright的追蹤log
        # self.context.tracing.start(screenshots=True, snapshots= True)

    def login(self, account, password):
        self.page.frame(name="fraInterface").click("[placeholder=\"輸入登錄名\"]")
        self.page.frame(name="fraInterface").fill(
            "[placeholder=\"輸入登錄名\"]", account)
        self.page.frame(name="fraInterface").click("[placeholder=\"輸入密碼\"]")
        self.page.frame(name="fraInterface").fill(
            "[placeholder=\"輸入密碼\"]", password)
        self.page.frame(name="fraInterface").click("text=登 錄")

    def logout(self):
        self.page.close()
        self.context.close()
        self.browser.close()


class noScan(LoginPage):
    # 繼承父類變數playwright: Playwright
    def __init__(self, playwright: Playwright):
        # Python3.x 和 Python2.x 的一个区别是: Python 3 可以使用直接使用 super().xxx 代替 super(Class, self).xxx :
        super().__init__(playwright)

    def SetTime(self, personinfo):
        self.page.frame(name="head").click("text=系統管理")
        self.page.frame(name="fraMenu").click("text=系統時間設定")
        # title = self.page.evaluate("document.title;")
        # print(title)
        self.page.frame(name="fraInterface").click("#TimeConf")
        # playwright執行javascript的方式
        self.page.frame(name="fraInterface").eval_on_selector(
            "#TimeConf", "el => el.removeAttribute('readonly')")
        # self.page.frame(name="fraInterface").evaluate('document.getElementById("TimeConf").removeAttribute("readonly");')
        date = personinfo["PolAppntDate"]
        self.page.frame(name="fraInterface").evaluate(
            f'document.getElementById("TimeConf").value = "{date}"')
        # # playwright要先確認跳出的alert
        self.page.on("dialog", lambda dialog: dialog.accept())
        self.page.frame(name="fraInterface").click("#SetTime")
        self.page.wait_for_timeout(2000)
        # self.page.wait_for_timeout(10000)
# -------------------------------------------------------------------------------------------------------------------------------------------------------
        # self.page.frame(name="head").click("text=系統管理")
        # self.page.frame(name="fraMenu").click("text=系統時間設定")
        # self.page.frame(name="fraInterface").click("#TimeConf")
        # self.page.frame(name="fraInterface").dispatch_event('//html/body/div/table/tbody/tr[3]/td[2]','click')

        # # playwright要先確認跳出的alert
        # self.page.on("dialog", lambda dialog: dialog.accept())
        # self.page.frame(name="fraInterface").click("#SetTime")
        # self.page.wait_for_timeout(1000)

    # def SetTime_Dec14(self):
    #     # 設定日期為12/14
    #     self.page.frame(name="head").click("text=系統管理")
    #     self.page.frame(name="fraMenu").click("text=系統時間設定")
    #     self.page.frame(name="fraInterface").click("#TimeConf")
    #     self.page.frame(name="fraInterface").dispatch_event('//html/body/div/table/tbody/tr[3]/td[3]','click')

    #     # playwright要先確認跳出的alert
    #     self.page.on("dialog", lambda dialog: dialog.accept())
    #     self.page.frame(name="fraInterface").click("#SetTime")
    #     self.page.wait_for_timeout(1000)

    # def SetTime_Dec20(self):
    #     # 設定日期為12/20
    #     self.page.frame(name="head").click("text=系統管理")
    #     self.page.frame(name="fraMenu").click("text=系統時間設定")
    #     self.page.frame(name="fraInterface").click("#TimeConf")
    #     self.page.wait_for_timeout(1000)
    #     self.page.frame(name="fraInterface").dispatch_event('//html/body/div/table/tbody/tr[4]/td[2]','click')

    #     # playwright要先確認跳出的alert
    #     self.page.on("dialog", lambda dialog: dialog.accept())
    #     self.page.frame(name="fraInterface").click("#SetTime")
    #     self.page.wait_for_timeout(1000)

    def Noscan_apply(self):
        # Click text=承保處理
        self.page.frame(name="head").click("text=承保處理")
        # Click text=個人保單
        self.page.frame(name="fraMenu").click("text=個人保單")
        # Click text=無掃描錄入
        self.page.frame(name="fraMenu").click("text=無掃描錄入")
        # # 輸入受理日期 SelfPoolQueryGrid3r0 不可使用受理日期申請保單，若要使用需調整營業日
        # self.page.frame(name="fraInterface").click("#SelfPoolQueryGrid3r0")
        # self.page.frame(name="fraInterface").fill("#SelfPoolQueryGrid3r0", personinfo["PolAppntDate"])
        # 申請保單按鈕
        self.page.frame(name="fraInterface").click("text=申請保單號碼")
        # 選擇第一個選項
        self.page.frame(name="fraInterface").check(
            "input[name=\"SelfPoolGridSel\"]")
        # Click text=開始錄入
        self.page.once("dialog", lambda dialog: dialog.accept())
        with self.page.expect_popup() as popup_info:
            self.page.frame(name="fraInterface").click("text=開始錄入")
            # self.page.screenshot(path="開始錄入.png")
        self.page2 = popup_info.value
        self.page2.set_viewport_size({"width": 1920, "height": 1080})

    def NoScan_query(self, personinfo):
        # Click text=承保處理
        self.page.frame(name="head").click("text=承保處理")
        # Click text=個人保單
        self.page.frame(name="fraMenu").click("text=個人保單")
        # Click text=無掃描錄入
        self.page.frame(name="fraMenu").click("text=無掃描錄入")
        # 受理日期(選填)
        self.page.frame(name="fraInterface").click(
            "input[name=\"SelfPoolQueryGrid3\"]")
        self.page.frame(name="fraInterface").fill(
            "input[name=\"SelfPoolQueryGrid3\"]", personinfo["PolAppntDate"])
        # 點擊査詢
        self.page.frame(name="fraInterface").click("input:has-text(\"査詢\")")

        # 選擇水池內  第一個選項 SelfPoolGridSel  第四個選項SelfPoolGridSel3
        self.page.frame(name="fraInterface").check(
            "input[name=\"SelfPoolGridSel\"]")
        # self.page.frame(name="fraInterface").check("#SelfPoolGridSel1")

        # Click text=開始錄入
        self.page.once("dialog", lambda dialog: dialog.accept())
        with self.page.expect_popup() as popup_info:
            self.page.frame(name="fraInterface").click("text=開始錄入")
        self.page2 = popup_info.value


class AddInfo(noScan):
    def __init__(self, playwright: Playwright):
        super().__init__(playwright)

# 個人資訊新增(傳統型)
    # def tranditional_Add(self, personinfo, product):
    #     test(需重寫幹娘娘)

# 個人資訊新增(投資型)
    def Add(self, personinfo):
        self.page2.touchscreen.tap(100, 100)
        # 填入要保書代號
        self.page2.frame(name="fraInterface").fill(
            "input[name=\"PolicyCode\"]", personinfo["PolicyCode"])
        self.page2.frame(name="fraInterface").click(
            "input[name=\"PolicyCode\"]")
        self.page2.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", personinfo["PolicyCode"])
        self.page2.frame(name="fraInterface").press(
            "input[name=\"PolicyCode\"]", "Enter")
        self.page2.frame(name="fraInterface").click(
            "input[name=\"ElectronicDoc\"]")

        # Click text=電話要約/要保日期 * 生效日期 * 受理日期 * >> :nth-match(img, 2)
        self.page2.frame(name="fraInterface").fill(
            "input[name=\"PolAppntDate\"]", personinfo["PolAppntDate"])

        self.page2.frame(name="fraInterface").fill(
            "input[name=\"CValiDate\"]", personinfo["CValiDate"])

        self.page2.frame(name="fraInterface").click(
            "input[name=\"OIUDelivery\"]")

        self.page2.frame(name="fraInterface").fill(
            "input[name=\"OIUDelivery\"]", personinfo["OIUDelivery"])

        self.page2.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", personinfo["OIUDelivery"])

        self.page2.frame(name="fraInterface").press(
            "input[name=\"OIUDelivery\"]", "Enter")

        self.page2.frame(name="fraInterface").click(
            "input[name=\"OnlineAttract\"]")

        # Fill input[name="OnlineAttract"]
        self.page2.frame(name="fraInterface").fill(
            "input[name=\"OnlineAttract\"]", personinfo["OnlineAttract"])

        self.page2.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", personinfo["OnlineAttract"])

        self.page2.frame(name="fraInterface").press(
            "input[name=\"OnlineAttract\"]", "Enter")

        # Click input[name="ElectronicDoc"]
        self.page2.frame(name="fraInterface").click(
            "input[name=\"ElectronicDoc\"]")

        # Fill input[name="ElectronicDoc"]
        self.page2.frame(name="fraInterface").fill(
            "input[name=\"ElectronicDoc\"]", personinfo["ElectronicDoc"])

        self.page2.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", personinfo["ElectronicDoc"])

        self.page2.frame(name="fraInterface").press(
            "input[name=\"ElectronicDoc\"]", "Enter")

        # 輸入錄音編號
        self.page2.frame(name="fraInterface").click("input[name=\"RecordNo\"]")
        self.page2.frame(name="fraInterface").fill(
            "input[name=\"RecordNo\"]", "123")

        # 輸入業務員資訊  0084309149田先生 0083302682楊＊芳
        self.page2.frame(name="fraInterface").click(
            "input[name=\"MultiAgentGridaddOne\"]")
        self.page2.frame(name="fraInterface").fill(
            "input[name=\"MultiAgentGrid1\"]", personinfo["Salesman"])
        self.page2.frame(name="fraInterface").click(
            "input[name=\"MultiAgentGrid1\"]")
        self.page2.frame(name="fraInterface").dblclick(
            "input[name=\"MultiAgentGrid1\"]")

        # 得到業務員業務員代號
        value1 = self.page2.frame(name="fraInterface").wait_for_selector(
            "//html/body/form/div[3]/div/table/tbody/tr/td/span/div[1]/div[5]/table/tbody/tr/td[7]/div/input")
        # value=value1.get_attribute('name')
        Salesman_value = value1.input_value()

        # 輸入姓名
        self.page2.frame(name="fraInterface").click(
            "input[name=\"AppntName\"]")
        self.page2.frame(name="fraInterface").fill(
            "input[name=\"AppntName\"]", personinfo["AppntName"])

        # 與被保險人關係
        self.page2.frame(name="fraInterface").click(
            "input[name=\"RelationToInsured\"]")

        self.page2.frame(name="fraInterface").fill(
            "input[name=\"RelationToInsured\"]", personinfo["RelationToInsured"])

        self.page2.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", personinfo["RelationToInsured"])

        self.page2.frame(name="fraInterface").press(
            "input[name=\"RelationToInsured\"]", "Enter")

        # 證件號碼
        self.page2.frame(name="fraInterface").click(
            "input[name=\"AppntIDNo\"]")

        self.page2.frame(name="fraInterface").fill(
            "input[name=\"AppntIDNo\"]", personinfo["AppntIDNo"])
        # 性別
        self.page2.frame(name="fraInterface").click("input[name=\"AppntSex\"]")

        self.page2.frame(name="fraInterface").fill(
            "input[name=\"AppntSex\"]", personinfo["AppntSex"])

        self.page2.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", personinfo["AppntSex"])
        # F:女生  M:男生

        self.page2.frame(name="fraInterface").press(
            "input[name=\"AppntSex\"]", "Enter")

        # 點擊證券號碼解除擋住問題
        self.page2.frame(name="fraInterface").click(
            "input[name=\"AppntIDNo\"]")
        # 國籍
        self.page2.frame(name="fraInterface").click(
            "input[name=\"AppntNativePlace\"]")

        self.page2.frame(name="fraInterface").fill(
            "input[name=\"AppntNativePlace\"]", "TW")

        self.page2.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", value='TW')

        self.page2.frame(name="fraInterface").press(
            "input[name=\"AppntNativePlace\"]", "Enter")
        # 生日
        self.page2.frame(name="fraInterface").fill(
            "input[name=\"AppntBirthday\"]", personinfo["AppntBirthday"])

        # 職業代碼
        self.page2.frame(name="fraInterface").click(
            "input[name=\"AppntOccupationCode\"]")

        self.page2.frame(name="fraInterface").fill(
            "input[name=\"AppntOccupationCode\"]", personinfo["AppntOccupationCode"])

        self.page2.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", personinfo["AppntOccupationCode"])

        # Press Enter
        self.page2.frame(name="fraInterface").press(
            "input[name=\"AppntOccupationCode\"]", "Enter")

        self.page2.frame(name="fraInterface").click("#codeselect")
        # 連絡電話
        self.page2.frame(name="fraInterface").click(
            "input[name=\"AppntCompanyPhoneCode\"]")
        # 身心障礙
        self.page2.frame(name="fraInterface").click(
            "input[name=\"DisabilityFlag\"]")

        self.page2.frame(name="fraInterface").fill(
            "input[name=\"DisabilityFlag\"]", personinfo["DisabilityFlag"])

        self.page2.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", personinfo["DisabilityFlag"])

        self.page2.frame(name="fraInterface").press(
            "input[name=\"DisabilityFlag\"]", "Enter")

        # email
        self.page2.frame(name="fraInterface").click(
            "input[name=\"AppntEMail\"]")
        self.page2.frame(name="fraInterface").fill(
            "input[name=\"AppntEMail\"]", "test@gmail.com")
        # 監護宣告
        self.page2.frame(name="fraInterface").click(
            "input[name=\"CustodyFlag\"]")

        self.page2.frame(name="fraInterface").fill(
            "input[name=\"CustodyFlag\"]", personinfo["CustodyFlag"])

        self.page2.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", personinfo["CustodyFlag"])

        self.page2.frame(name="fraInterface").press(
            "input[name=\"CustodyFlag\"]", "Enter")

        # 婚姻狀況
        self.page2.frame(name="fraInterface").click(
            "input[name=\"AppntMarriage\"]")

        self.page2.frame(name="fraInterface").fill(
            "input[name=\"AppntMarriage\"]", personinfo["AppntMarriage"])

        self.page2.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", personinfo["AppntMarriage"])

        self.page2.frame(name="fraInterface").press(
            "input[name=\"AppntMarriage\"]", "Enter")

        # email
        self.page2.frame(name="fraInterface").click(
            "input[name=\"AppntEMail\"]")
        # 行動電話
        self.page2.frame(name="fraInterface").click(
            "input[name=\"AppntMobile\"]")
        self.page2.frame(name="fraInterface").fill(
            "input[name=\"AppntMobile\"]", "0987654321")
        # 住所
        self.page2.frame(name="fraInterface").fill(
            "input[name=\"HomeCityName\"]", personinfo["HomeCityName"])
        self.page2.frame(name="fraInterface").fill(
            "input[name=\"HomeCountyName\"]", personinfo["HomeCountyName"])
        with self.page2.expect_popup() as popup_info:
            self.page2.frame(name="fraInterface").dblclick(
                "[placeholder=\"請雙擊\"]")
        self.page3 = popup_info.value

        self.page3.click("input[name=\"Street\"]")

        self.page3.fill("input[name=\"Street\"]", "測試")
        # Click text=保存
        self.page3.click("text=保存")
        self.page3.close()
        # 營業項目
        self.page2.frame(name="fraInterface").fill(
            "input[name=\"BusinessProject\"]", "1")
        # 服務單位
        self.page2.frame(name="fraInterface").fill(
            "input[name=\"ServiceUnit\"]", "1")
        # 職位
        self.page2.frame(name="fraInterface").fill(
            "input[name=\"AppntPosition\"]", "1")
        # # 保存按鈕
        self.page2.frame(name="fraInterface").click("#addbutton")
        # 要保人為本人之勾選框
        self.page2.frame(name="fraInterface").click(
            "input[name=\"SamePersonFlag\"]")
        # 保存

        # ---------------
        self.page2.frame(name="fraInterface").click("#divAddInsuredButtun")

        # 繳費資訊
        # 首期
        self.page2.frame(name="fraInterface").dblclick(
            "input[name=\"PayMode2\"]")
        self.page2.frame(name="fraInterface").fill(
            "input[name=\"PayMode2\"]", personinfo["PayMode2"])

        self.page2.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", personinfo["PayMode2"])

        self.page2.frame(name="fraInterface").press(
            "input[name=\"PayMode2\"]", "Enter")

        # 續期
        self.page2.frame(name="fraInterface").dblclick(
            "input[name=\"SecPayMode2\"]")
        self.page2.frame(name="fraInterface").fill(
            "input[name=\"SecPayMode2\"]", personinfo["SecPayMode2"])

        self.page2.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", personinfo["SecPayMode2"])

        self.page2.frame(name="fraInterface").press(
            "input[name=\"SecPayMode2\"]", "Enter")

        # 繳款人身份
        self.page2.frame(name="fraInterface").dblclick(
            "input[name=\"StatusOfPayer\"]")
        self.page2.frame(name="fraInterface").fill(
            "input[name=\"StatusOfPayer\"]", personinfo["StatusOfPayer"])

        self.page2.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", personinfo["StatusOfPayer"])

        self.page2.frame(name="fraInterface").press(
            "input[name=\"StatusOfPayer\"]", "Enter")

        # 保存
        # 因保存在頁面太多重複，此語法為點擊頁面內第三個保存
        self.page2.frame(name="fraInterface").click(
            ":nth-match(:text('保存'), 3)")

        print("業務員代號:"+Salesman_value)
        return Salesman_value

# ------------------------------------------------------------------------------------
# 此處是因為個人的"職業代碼"在保存後，進行商品錄入時會出現錯誤，因此必須先在此處進行商品錄入後，進入商品錄入(product add)的方法時才能進行修改解決問題
        # self.page2.wait_for_timeout(3000)
        # self.page2.frame(name="fraInterface").click("text=商品資訊錄入")
        # self.page2.wait_for_timeout(2000)

        # self.page2.frame(name="fraInterface").dblclick("input[name=\"RiskCode\"]")
        # self.page2.frame(name="fraInterface").fill("input[name=\"RiskCode\"]", product["productId"])

        # self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["productId"])

        # self.page2.frame(name="fraInterface").press("input[name=\"RiskCode\"]", "Enter")

        # # Click text=進入商品錄入畫面
        # self.page2.frame(name="fraInterface").click("input[name=\"back\"]")
        # self.page2.wait_for_timeout(2000)

        # self.page2.frame(name="fraInterface").click("input[name=\"Prem\"]")
        # self.page2.frame(name="fraInterface").fill("input[name=\"Prem\"]", product["price"])
        # # 年金給付開始年齡
        # self.page2.frame(name="fraInterface").click("input[name=\"GetYear\"]")
        # self.page2.frame(name="fraInterface").fill("input[name=\"GetYear\"]", product["age"])
        # # 年金給付方式
        # self.page2.frame(name="fraInterface").dblclick("input[name=\"GetIntv\"]")

        # self.page2.frame(name="fraInterface").fill("input[name=\"GetIntv\"]", "0")

        # # 給付方式
        # self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",value='0')

        # self.page2.frame(name="fraInterface").press("input[name=\"GetIntv\"]", "Enter")

        # # 保存
        # self.page2.frame(name="fraInterface").click("text=保 存")

        # # 上一步
        # self.page2.frame(name="fraInterface").click("#riskbutton1")
# ------------------------------------------------------------------------------------

# 投資型商品錄入(NIFA1603)
    def ProductAdd(self, product):
        self.page2.touchscreen.tap(100, 100)
        # 保存結束後要再按一次修改，不然投資標的的新增會失敗
        self.page2.frame(name="fraInterface").click("#updatebutton")
        self.page2.wait_for_timeout(5000)
        self.page2.frame(name="fraInterface").click("text=商品資訊錄入")
        self.page2.wait_for_timeout(2000)

        # 輸入投標的
        self.page2.frame(name="fraInterface").dblclick(
            "input[name=\"RiskCode\"]")
        self.page2.frame(name="fraInterface").fill(
            "input[name=\"RiskCode\"]", product["productId"])
        self.page2.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", product["productId"])
        self.page2.frame(name="fraInterface").press(
            "input[name=\"RiskCode\"]", "Enter")

        # 進入商品錄入畫面
        self.page2.frame(name="fraInterface").click("input[name=\"back\"]")
        self.page2.wait_for_timeout(2000)

        # 輸入保額
        self.page2.frame(name="fraInterface").click("input[name=\"Prem\"]")
        self.page2.frame(name="fraInterface").fill(
            "input[name=\"Prem\"]", product["price"])

        # 年金給付開始年齡
        self.page2.frame(name="fraInterface").click("input[name=\"GetYear\"]")
        self.page2.frame(name="fraInterface").fill(
            "input[name=\"GetYear\"]", product["age"])

        # 年金給付方式
        self.page2.frame(name="fraInterface").dblclick(
            "input[name=\"GetIntv\"]")
        self.page2.frame(name="fraInterface").fill(
            "input[name=\"GetIntv\"]", "0")

        # 給付方式
        self.page2.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", value='0')
        self.page2.frame(name="fraInterface").press(
            "input[name=\"GetIntv\"]", "Enter")

        # 保存
        self.page2.frame(name="fraInterface").click("text=保 存")
        self.page2.wait_for_timeout(2000)

        # 輸入投資目標
        self.page2.frame(name="fraInterface").dblclick(
            "input[name=\"InvestPlanRate1\"]")
        self.page2.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", value="E002")

        # 輸入投資%數
        self.page2.frame(name="fraInterface").press(
            "input[name=\"InvestPlanRate1\"]", "Enter")
        self.page2.frame(name="fraInterface").fill(
            "input[name=\"InvestPlanRate5\"]", product["InvestPlanRate"])
        # 添加
        self.page2.frame(name="fraInterface").click(
            "input[name=\"submitFormButton\"]")
        # 上一步
        self.page2.frame(name="fraInterface").click("#riskbutton1")
        # self.page2.screenshot(path="test4.png")
# 改為使用ProcessControl料夾內的Tranditional.py
# # 傳統型商品錄入(NTWL0202)
#     def tranditional_productAdd_1113_NTWL0202(self,product):
#         self.page2.touchscreen.tap(100,100)
#         # 保存結束後要再按一次修改，不然投資標的的新增會失敗
#         self.page2.frame(name="fraInterface").click("#updatebutton")
#         self.page2.wait_for_timeout(5000)
#         self.page2.frame(name="fraInterface").click("text=商品資訊錄入")
#         self.page2.wait_for_timeout(2000)

#         self.page2.frame(name="fraInterface").dblclick("input[name=\"RiskCode\"]")
#         self.page2.frame(name="fraInterface").fill("input[name=\"RiskCode\"]", product["productId"])

#         self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["productId"])

#         self.page2.frame(name="fraInterface").press("input[name=\"RiskCode\"]", "Enter")

#         # Click text=進入商品錄入畫面
#         self.page2.frame(name="fraInterface").click("input[name=\"back\"]")
#         self.page2.wait_for_timeout(2000)

#         # 保額
#         self.page2.frame(name="fraInterface").click("input[name=\"Amnt\"]")
#         self.page2.frame(name="fraInterface").fill("input[name=\"Amnt\"]", product["price"])

#         # 繳別
#         self.page2.frame(name="fraInterface").dblclick("input[name=\"PayIntv\"]")
#         self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["PayIntv"])
#         self.page2.frame(name="fraInterface").press("input[name=\"PayIntv\"]", "Enter")

#         # 繳費期間
#         self.page2.frame(name="fraInterface").dblclick("input[name=\"PayEndYear\"]")
#         self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["PayEndYear"])
#         self.page2.frame(name="fraInterface").press("input[name=\"PayEndYear\"]", "Enter")


#         # 自動墊繳
#         self.page2.frame(name="fraInterface").dblclick("input[name=\"AutoPayFlag\"]")
#         self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",product["AutoPayFlag"])
#         self.page2.frame(name="fraInterface").press("input[name=\"AutoPayFlag\"]", "Enter")

#         # 保存
#         self.page2.frame(name="fraInterface").click("text=保 存")
#         self.page2.wait_for_timeout(2000)

#         # 上一步
#         self.page2.frame(name="fraInterface").click("#riskbutton1")


# 業務員招攬報告書

    def SalesmanReport(self, SaleReport):
        self.page2.touchscreen.tap(100, 100)
        self.page2.once("dialog", lambda dialog: dialog.accept())
        with self.page2.expect_popup() as popup_info:
            self.page2.frame(name="fraInterface").click("#AgentImpart")
        self.page5 = popup_info.value
        self.page5.wait_for_load_state()
        print(self.page5.title())
        self.page5.screenshot(path="業務員招攬報告書.png")

        # 選擇主約險種
        self.page5.frame(name="fraInterface").dblclick(
            "input[name=\"MainPolNo\"]")
        self.page5.frame(name="fraInterface").fill(
            "input[name=\"MainPolNo\"]", SaleReport["MainPolNo"])
        self.page5.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", SaleReport["MainPolNo"])
        self.page5.frame(name="fraInterface").press(
            "input[name=\"MainPolNo\"]", "Enter")

        # 輸入要保人姓名
        self.page5.frame(name="fraInterface").fill(
            "input[name=\"AppntNo\"]", SaleReport["AppntNo"])

        # 是否立即繳款
        self.page5.frame(name="fraInterface").dblclick(
            "input[name=\"IsNowPay\"]")
        self.page5.frame(name="fraInterface").fill(
            "input[name=\"IsNowPay\"]", "N")
        self.page5.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", SaleReport["IsNowPay"])

        self.page5.frame(name="fraInterface").press(
            "input[name=\"IsNowPay\"]", "Enter")
        # 幣別
        self.page5.frame(name="fraInterface").dblclick(
            "input[name=\"Currency\"]")
        self.page5.frame(name="fraInterface").fill(
            "input[name=\"Currency\"]", SaleReport["Currency"])
        self.page5.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", SaleReport["Currency"])
        self.page5.frame(name="fraInterface").press(
            "input[name=\"Currency\"]", "Enter")

        # 親友保戶介紹
        self.page5.frame(name="fraInterface").dblclick("#AgentImpartsGrid5r3")
        # self.page5.frame(name="fraInterface").fill("#AgentImpartsGrid5r3", "Y")
        self.page5.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", SaleReport["AgentImpartsGrid5r3"])
        self.page5.frame(name="fraInterface").press(
            "#AgentImpartsGrid5r3", "Enter")
        # 保存
        self.page5.frame(name="fraInterface").click("#clickSave")
        self.page5.frame(name="fraInterface").click("text=返回")

# 受益人資訊錄入(投資型)
    def BenefitPeople(self, BenefitInfo):
        # time.sleep(2)
        self.page2.screenshot(path="受益人資料錄入2.png")
        self.page2.touchscreen.tap(100, 100)
        # self.page2.screenshot(path="受益人資料錄入.png")
        self.page2.once("dialog", lambda dialog: dialog.accept())
        with self.page2.expect_popup() as popup_info:
            self.page2.frame(name="fraInterface").click("#LBnfButton")
        self.page6 = popup_info.value
        self.page6.set_viewport_size({"width": 1920, "height": 1080})

        # # 受益人類別
        # BnfType 身故受益人
        self.page6.touchscreen.tap(100, 100)
        self.page6.frame(name="fraInterface").fill(
            "#BnfType", BenefitInfo["BnfType1"])
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
        self.page6.frame(name="fraInterface").fill(
            "#BnfGrade", BenefitInfo["BnfGrade"])
        self.page6.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", BenefitInfo["BnfGrade"])
        self.page6.wait_for_timeout(3000)

        self.page6.frame(name="fraInterface").press("#BnfGrade", "Enter")
        # 100%
        self.page6.frame(name="fraInterface").fill(
            "#BnfLot", BenefitInfo["BnfLot"])
        # 保存
        self.page6.frame(name="fraInterface").click("#save")
        self.page6.once("dialog", lambda dialog: dialog.dismiss())

        # # 受益人類別
        # BnfType 生存金受益人受益人
        self.page6.frame(name="fraInterface").dblclick("#BnfType")
        self.page6.frame(name="fraInterface").fill("#BnfType", "3")
        self.page6.frame(name="fraInterface").click("option[value=\"3\"]")
        self.page6.wait_for_timeout(3000)

        # # 同要保人 check
        # isAppnt
        self.page6.frame(name="fraInterface").check("#isAppnt")
        # # 受益人順位
        # BnfGrade 第一順位
        self.page6.frame(name="fraInterface").dblclick("#BnfGrade")
        self.page6.frame(name="fraInterface").fill("#BnfGrade", "01")
        self.page6.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", value='01')
        self.page6.wait_for_timeout(3000)

        self.page6.frame(name="fraInterface").press("#BnfGrade", "Enter")
        # 100%
        self.page6.frame(name="fraInterface").fill("#BnfLot", "100")
        self.page6.frame(name="fraInterface").fill(
            "input[name=\"BnfPhone\"]", "0987654321")
        # 保存
        self.page6.frame(name="fraInterface").click("#save")
        # self.page6.once("dialog", lambda dialog: dialog.dismiss())

        # # 受益人類別
        # BnfType 年金金受益人受益人
        self.page6.frame(name="fraInterface").dblclick("#BnfType")
        self.page6.frame(name="fraInterface").fill("#BnfType", "5")
        self.page6.frame(name="fraInterface").click("option[value=\"5\"]")
        self.page6.wait_for_timeout(3000)

        # # 法定 check
        # isAppnt
        self.page6.frame(name="fraInterface").check("#isSameLaw")
        # # 受益人順位
        # BnfGrade 第一順位
        self.page6.frame(name="fraInterface").dblclick("#BnfGrade")
        self.page6.frame(name="fraInterface").fill("#BnfGrade", "01")
        self.page6.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", value='01')
        self.page6.wait_for_timeout(3000)

        self.page6.frame(name="fraInterface").press("#BnfGrade", "Enter")
        # 100%
        self.page6.frame(name="fraInterface").fill("#BnfLot", "100")

        # 保存
        self.page6.frame(name="fraInterface").click("#save")
        self.page6.once("dialog", lambda dialog: dialog.dismiss())
        self.page6.frame(name="fraInterface").click("text=返回")

# 受益人資訊錄入(傳統型)
    def tranditional_BenefitPeople(self, BenefitInfo):
        # time.sleep(2)
        self.page2.screenshot(path="受益人資料錄入2.png")
        self.page2.touchscreen.tap(100, 100)
        # self.page2.screenshot(path="受益人資料錄入.png")
        self.page2.once("dialog", lambda dialog: dialog.accept())
        with self.page2.expect_popup() as popup_info:
            self.page2.frame(name="fraInterface").click("#LBnfButton")
        self.page6 = popup_info.value
        self.page6.set_viewport_size({"width": 1920, "height": 1080})

        # # 受益人類別
        # BnfType 身故受益人
        self.page6.touchscreen.tap(100, 100)
        self.page6.frame(name="fraInterface").fill(
            "#BnfType", BenefitInfo["BnfType1"])
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
        self.page6.frame(name="fraInterface").fill(
            "#BnfGrade", BenefitInfo["BnfGrade"])
        self.page6.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", BenefitInfo["BnfGrade"])
        self.page6.wait_for_timeout(3000)
        self.page6.frame(name="fraInterface").press("#BnfGrade", "Enter")
        # 100%
        self.page6.frame(name="fraInterface").fill(
            "#BnfLot", BenefitInfo["BnfLot"])
        # 保存
        self.page6.frame(name="fraInterface").click("#save")
        self.page6.once("dialog", lambda dialog: dialog.dismiss())

        # # 受益人類別
        # 祝壽受益人
        self.page6.frame(name="fraInterface").dblclick("#BnfType")
        # 選擇祝壽受益人
        self.page6.frame(name="fraInterface").fill("#BnfType", "2")
        self.page6.frame(name="fraInterface").click("option[value=\"2\"]")
        self.page6.wait_for_timeout(3000)

        # # 同要保人 check
        # isAppnt
        self.page6.frame(name="fraInterface").check("#isAppnt")

        # # 受益人順位
        # BnfGrade 第一順位
        self.page6.frame(name="fraInterface").dblclick("#BnfGrade")
        self.page6.frame(name="fraInterface").fill("#BnfGrade", "01")
        self.page6.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", value='01')
        self.page6.wait_for_timeout(3000)
        self.page6.frame(name="fraInterface").press("#BnfGrade", "Enter")

        # 100%
        self.page6.frame(name="fraInterface").fill("#BnfLot", "100")
        self.page6.frame(name="fraInterface").fill(
            "input[name=\"BnfPhone\"]", "0987654321")

        # 保存
        self.page6.frame(name="fraInterface").click("#save")
        self.page6.once("dialog", lambda dialog: dialog.dismiss())
        self.page6.frame(name="fraInterface").click("text=返回")

    # 重要事項告知書
    def ImportInf(self):
        self.page2.touchscreen.tap(100, 100)
        # self.page2.screenshot(path="重要事項告知書.png")
        self.page2.once("dialog", lambda dialog: dialog.accept())
        with self.page2.expect_popup() as popup_info:
            self.page2.frame(name="fraInterface").click("#ImpInf")
        self.page7 = popup_info.value
        self.page7.set_viewport_size({"width": 1920, "height": 1080})
        # self.page7.screenshot(path="重要事項告知書2.png")
        self.page7.touchscreen.tap(100, 100)
        print(self.page7.title())

        # (全部打勾)
        self.page7.frame(name="fraInterface").dispatch_event(
            "//html/body/form/div[3]/div/table/tbody/tr/td/span/div[1]/div[3]/div/table/thead/tr/th[1]/div/input", "click")
        # self.page7.frame(name="fraInterface").check("#ImportInformGridChk0")
        # # ImportInform2GridChk0
        # self.page7.frame(name="fraInterface").check("#ImportInformGridChk1")
        # self.page7.frame(name="fraInterface").check("#ImportInformGridChk2")
        self.page7.frame(name="fraInterface").click("#save")
        self.page7.once("dialog", lambda dialog: dialog.dismiss())
        self.page7.wait_for_timeout(2000)

        self.page7.frame(name="fraInterface").click("text=返回")

# 投資屬性分析問券
    def InvestReport(self, personinfo):
        self.page2.touchscreen.tap(100, 100)
        self.page2.once("dialog", lambda dialog: dialog.accept())
        with self.page2.expect_popup() as popup_info:
            self.page2.frame(name="fraInterface").click("#InAttQue")
        self.page8 = popup_info.value
        self.page8.set_viewport_size({"width": 1920, "height": 1080})
        # self.page8.touchscreen.tap(100,100)
        self.page8.screenshot(path="投資屬性分析問卷.png")
        print(self.page8.title())

        if(20 < int(personinfo['age']) <= 44):
            # 20-44
            print("20歲以上")
            self.page8.frame(name="fraInterface").check("#AgeGridSel1")
            self.page8.frame(name="fraInterface").check("#JobGridSel0")
        elif(45 < int(personinfo['age']) <= 65):
            # 45-65
            print("45歲以上")
            self.page8.frame(name="fraInterface").check("#AgeGridSel2")
            self.page8.frame(name="fraInterface").check("#JobGridSel0")
        elif(66 < int(personinfo['age']) <= 69):
            # 66-69
            print("66歲以上")
            self.page8.frame(name="fraInterface").check("#AgeGridSel3")
            self.page8.frame(name="fraInterface").check("#JobGridSel0")
        elif(int(personinfo['age']) > 70):
            # >70
            print("70歲以上")
            self.page8.frame(name="fraInterface").check("#AgeGridSel4")
            self.page8.frame(name="fraInterface").check("#JobGridSel0")
        elif(0 < int(personinfo['age'] <= 20)):
            # 0-20歲
            print("20歲以下")

        self.page8.frame(name="fraInterface").check("#ExeGridChk2")
        self.page8.frame(name="fraInterface").check("#ExeGridChk4")
        self.page8.frame(name="fraInterface").check("#ExeGridChk6")

        self.page8.frame(name="fraInterface").check("#RiskGridSel0")

        if(20 < int(personinfo['age']) <= 44):
            self.page8.frame(name="fraInterface").fill("#Total", value='45')
        elif(45 < int(personinfo['age']) <= 65):
            self.page8.frame(name="fraInterface").fill("#Total", value='47')
        elif(66 < int(personinfo['age']) <= 69):
            self.page8.frame(name="fraInterface").fill("#Total", value='43')
        elif(int(personinfo['age']) > 70):
            self.page8.frame(name="fraInterface").fill("#Total", value='42')

        self.page8.screenshot(path="投資屬性分析問卷.png")

        # 同要保書電話
        self.page8.frame(name="fraInterface").dblclick("#CallPhoneFlag")
        self.page8.frame(name="fraInterface").fill("#CallPhoneFlag", "1")
        self.page8.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", value='1')
        self.page8.wait_for_timeout(3000)
        self.page8.frame(name="fraInterface").press("#CallPhoneFlag", "Enter")
        # 投資人須知]、[保險商品說明書]交付確認書
        self.page8.frame(name="fraInterface").check("#Confirm1GridChk0")
        self.page8.frame(name="fraInterface").check("#Confirm1GridChk1")
        self.page8.frame(name="fraInterface").check("#Confirm1GridChk2")
        self.page8.frame(name="fraInterface").check("#Confirm1GridChk3")

        self.page8.frame(name="fraInterface").check("#Confirm2GridChk0")
        self.page8.frame(name="fraInterface").check("#Confirm2GridChk1")

        self.page8.frame(name="fraInterface").click("#save")
        self.page8.wait_for_timeout(2000)
        # self.page8.screenshot(path="投資屬性分析問卷2.png")
        self.page8.frame(name="fraInterface").click("text=返回")
# ---舊的投資屬性分析---------------------------------------------------------------------
        # self.page2.touchscreen.tap(100,100)
        # self.page2.once("dialog", lambda dialog: dialog.accept())
        # with self.page2.expect_popup() as popup_info:
        #     self.page2.frame(name="fraInterface").click("#InAttQue")
        # self.page8 = popup_info.value
        # self.page8.set_viewport_size({"width": 1920, "height": 1080})
        # # self.page8.touchscreen.tap(100,100)
        # self.page8.screenshot(path="投資屬性分析問卷.png")
        # print(self.page8.title())

        # self.page8.frame(name="fraInterface").check("#AgeGridSel1")
        # self.page8.frame(name="fraInterface").check("#JobGridSel0")

        # self.page8.frame(name="fraInterface").check("#ExeGridChk2")
        # self.page8.frame(name="fraInterface").check("#ExeGridChk4")
        # self.page8.frame(name="fraInterface").check("#ExeGridChk6")

        # self.page8.frame(name="fraInterface").check("#RiskGridSel0")
        # self.page8.frame(name="fraInterface").fill("#Total",value='45')

        # # 同要保書電話
        # self.page8.frame(name="fraInterface").dblclick("#CallPhoneFlag")
        # self.page8.frame(name="fraInterface").fill("#CallPhoneFlag", "1")
        # self.page8.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",value='1')
        # time.sleep(2)
        # self.page8.frame(name="fraInterface").press("#CallPhoneFlag", "Enter")
        # # 投資人須知]、[保險商品說明書]交付確認書
        # self.page8.frame(name="fraInterface").check("#Confirm1GridChk0")
        # self.page8.frame(name="fraInterface").check("#Confirm1GridChk1")
        # self.page8.frame(name="fraInterface").check("#Confirm1GridChk2")
        # self.page8.frame(name="fraInterface").check("#Confirm1GridChk3")

        # self.page8.frame(name="fraInterface").check("#Confirm2GridChk0")
        # self.page8.frame(name="fraInterface").check("#Confirm2GridChk1")

        # self.page8.frame(name="fraInterface").click("#save")
        # self.page8.wait_for_timeout(2000)
        # # self.page8.screenshot(path="投資屬性分析問卷2.png")
        # self.page8.frame(name="fraInterface").click("text=返回")
# -----------------------------------------------------------------------------------------------

    def FATCA(self, personinfo):
        self.page2.touchscreen.tap(100, 100)
        self.page2.once("dialog", lambda dialog: dialog.accept())
        with self.page2.expect_popup() as popup_info:
            self.page2.frame(name="fraInterface").click("#riskbutton53")
        self.page9 = popup_info.value
        self.page9.set_viewport_size({"width": 1920, "height": 1080})
        print(self.page9.title())
        # self.page8.touchscreen.tap(100,100)
        # self.page9.screenshot(path="FATCA.png")
        self.page9.touchscreen.tap(100, 100)
        self.page9.frame(name="fraInterface").click("#BirthPlace")
        self.page9.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", value='01')
        time.sleep(2)
        self.page9.frame(name="fraInterface").click(
            "select[name=\"codeselect\"]")

        self.page9.frame(name="fraInterface").click("#FatcaType")
        self.page9.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", value='01')
        time.sleep(2)
        self.page9.frame(name="fraInterface").click(
            "select[name=\"codeselect\"]")
        self.page9.frame(name="fraInterface").fill(
            "#FatcaStatementDate", personinfo["PolAppntDate"])
        self.page9.frame(name="fraInterface").click("#save")
        self.page9.wait_for_timeout(2000)
        # self.page9.screenshot(path="FATCA.png")
        self.page9.frame(name="fraInterface").click("text=返回")

    def CRS(self, personinfo):
        self.page2.touchscreen.tap(100, 100)
        self.page2.once("dialog", lambda dialog: dialog.accept())
        with self.page2.expect_popup() as popup_info:
            self.page2.frame(name="fraInterface").click("#riskbutton52")
        self.page10 = popup_info.value
        self.page10.set_viewport_size({"width": 1920, "height": 1080})
        print(self.page10.title())
        # self.page8.touchscreen.tap(100,100)
        # self.page9.screenshot(path="FATCA.png")
        self.page10.touchscreen.tap(100, 100)

        self.page10.frame(name="fraInterface").click("#CRSType")
        self.page10.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", value='N')
        time.sleep(2)
        self.page10.frame(name="fraInterface").click(
            "select[name=\"codeselect\"]")

        self.page10.frame(name="fraInterface").fill(
            "#NoticeDate", personinfo["PolAppntDate"])
        # self.page10.screenshot(path="CRS2.png")
        # 保存
        self.page10.frame(name="fraInterface").click("#SCinsertButton")

        # 修改
        # self.page10.frame(name="fraInterface").click("#SCupdateButton")

# 合約其他資訊(投資型)
    def invest_contract_other_info(self):
        self.page2.touchscreen.tap(100, 100)
        # 保單製單類型
        self.page2.frame(name="fraInterface").click("#PolicyPreparationType")
        self.page2.frame(name="fraInterface").fill(
            "#PolicyPreparationType", "1")
        self.page2.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", value='1')
        self.page2.wait_for_timeout(3000)

        self.page2.frame(name="fraInterface").press(
            "#PolicyPreparationType", "Enter")
        # # 結匯授權書(需加入excel當作變數)
        # self.page2.frame(name="fraInterface").click("#SettlementAuthorization")
        # self.page2.frame(name="fraInterface").fill("#SettlementAuthorization", "Y")
        # self.page2.frame(name="fraInterface").select_option("select[name=\"codeselect\"]",value='Y')
        # self.page2.wait_for_timeout(3000)
        # self.page2.frame(name="fraInterface").press("#SettlementAuthorization", "Enter")

        # 增值回饋分享金、健康回饋獎勵金、帳戶價值通知方式
        self.page2.frame(name="fraInterface").click("#Notification")
        self.page2.frame(name="fraInterface").fill("#Notification", "1")
        self.page2.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", value='1')
        self.page2.wait_for_timeout(3000)
        self.page2.frame(name="fraInterface").press("#Notification", "Enter")

        # 結匯授權書(可要可不要)
        self.page2.frame(name="fraInterface").click("#SettlementAuthorization")
        self.page2.frame(name="fraInterface").fill(
            "#SettlementAuthorization", "Y")
        self.page2.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", value='Y')
        self.page2.wait_for_timeout(3000)
        self.page2.frame(name="fraInterface").press(
            "#SettlementAuthorization", "Enter")

        self.page2.frame(name="fraInterface").click(
            ":nth-match(:text('保存'), 4)")
        self.page2.wait_for_timeout(3000)
        # 錄入完畢
        self.page2.frame(name="fraInterface").click("#riskbutton2")
        # self.context.tracing.stop(path = "trace.zip")

# 合約其他資訊(傳統型)
    def tranditional_contract_other_info(self, product):
        self.page2.touchscreen.tap(100, 100)
        # 保單製單類型
        self.page2.frame(name="fraInterface").click("#PolicyPreparationType")
        self.page2.frame(name="fraInterface").fill(
            "#PolicyPreparationType", "1")
        self.page2.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", value='1')
        self.page2.wait_for_timeout(1000)
        self.page2.frame(name="fraInterface").press(
            "#PolicyPreparationType", "Enter")

        # 增值回饋分享金、健康回饋獎勵金、帳戶價值通知方式
        self.page2.frame(name="fraInterface").click("#Notification")
        self.page2.frame(name="fraInterface").fill("#Notification", "1")
        self.page2.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", value='1')
        self.page2.wait_for_timeout(1000)
        self.page2.frame(name="fraInterface").press("#Notification", "Enter")

        # 以外幣收付之非投資型人身保險匯率風險說明書(不等於才會進去)
        if(product["productId"] != 'NTWL0105' and product["productId"] != 'NIFA0802'):
            self.page2.frame(name="fraInterface").click("#WhetherStatement")
            self.page2.frame(name="fraInterface").fill(
                "#WhetherStatement", "Y")
            self.page2.frame(name="fraInterface").select_option(
                "select[name=\"codeselect\"]", value='Y')
            self.page2.wait_for_timeout(1000)
            self.page2.frame(name="fraInterface").press(
                "#WhetherStatement", "Enter")

        self.page2.frame(name="fraInterface").click(
            ":nth-match(:text('保存'), 4)")
        self.page2.wait_for_timeout(1000)
        # 錄入完畢
        self.page2.frame(name="fraInterface").click("#riskbutton2")
        # self.context.tracing.stop(path = "trace.zip")

# 共同行銷及特定目的外搜集、處理、利用個人資料之聲明同意書
    def commonpurpose(self):
        self.page2.touchscreen.tap(100, 100)
        self.page2.once("dialog", lambda dialog: dialog.accept())
        with self.page2.expect_popup() as popup_info:
            self.page2.frame(name="fraInterface").click("#riskbutton58")
        self.page12 = popup_info.value
        self.page12.set_viewport_size({"width": 1920, "height": 1080})
        self.page12.wait_for_timeout(3000)
        print(self.page12.title())

        self.page12.frame(name="fraInterface").click("#EncloseStatement")
        self.page12.frame(name="fraInterface").fill("#EncloseStatement", "N")
        self.page12.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", value='N')
        # self.page12.wait_for_timeout(2000)
        self.page12.frame(name="fraInterface").press(
            "#EncloseStatement", "Enter")

        self.page12.frame(name="fraInterface").click("#IFAgree")
        self.page12.frame(name="fraInterface").fill("#IFAgree", "N")
        self.page12.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", value='N')
        # self.page12.wait_for_timeout(2000)
        self.page12.frame(name="fraInterface").press("#IFAgree", "Enter")

        self.page12.frame(name="fraInterface").click("#cEncloseStatement")
        self.page12.frame(name="fraInterface").fill("#cEncloseStatement", "N")
        self.page12.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", value='N')
        # self.page12.wait_for_timeout(2000)
        self.page12.frame(name="fraInterface").press(
            "#cEncloseStatement", "Enter")

        self.page12.frame(name="fraInterface").click("#cIFAgree")
        self.page12.frame(name="fraInterface").fill("#cIFAgree", "N")
        self.page12.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", value='N')
        # self.page12.wait_for_timeout(2000)
        self.page12.frame(name="fraInterface").press("#cIFAgree", "Enter")

        self.page12.frame(name="fraInterface").click("text=保存")
        self.page12.wait_for_timeout(2000)
        self.page12.close()


# 法定代理人(15歲)

    def Legal_representative(self, Agent):
        self.page2.touchscreen.tap(100, 100)
        self.page2.once("dialog", lambda dialog: dialog.accept())
        with self.page2.expect_popup() as popup_info:
            self.page2.frame(name="fraInterface").click("#riskbutton50")
        self.page13 = popup_info.value
        self.page13.set_viewport_size({"width": 1920, "height": 1080})
        self.page13.wait_for_timeout(3000)
        print(self.page13.title())

        # 姓名
        self.page13.frame(name="fraInterface").fill(
            "#LegalName", Agent["LegalName"])

        # 出生日期
        self.page13.frame(name="fraInterface").fill(
            "#LegalBrirthDay", Agent["LegalBrirthDay"])

        # 與未成年人關係
        self.page13.frame(name="fraInterface").click("#RelationTolegal")
        self.page13.frame(name="fraInterface").fill(
            "#RelationTolegal", Agent["RelationTolegal"])
        self.page13.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", Agent["RelationTolegal"])
        self.page13.wait_for_timeout(2000)
        self.page13.frame(name="fraInterface").press(
            "#RelationTolegal", "Enter")

        # 法定代理人類型
        self.page13.frame(name="fraInterface").click("#GuardianType")
        self.page13.frame(name="fraInterface").fill(
            "#GuardianType", Agent["GuardianType"])
        self.page13.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", Agent["GuardianType"])
        self.page13.wait_for_timeout(2000)
        self.page13.frame(name="fraInterface").press("#GuardianType", "Enter")

        # 證件類型
        self.page13.frame(name="fraInterface").click("#LegalIDType")
        self.page13.frame(name="fraInterface").fill(
            "#LegalIDType", Agent["LegalIDType"])
        self.page13.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", Agent["LegalIDType"])
        self.page13.wait_for_timeout(2000)
        self.page13.frame(name="fraInterface").press("#LegalIDType", "Enter")

        # 證件號碼
        self.page13.frame(name="fraInterface").fill(
            "#LegalIDNo", Agent["LegalIDNo"])

        # 國籍
        self.page13.frame(name="fraInterface").click("#LegalNativePlace")
        self.page13.frame(name="fraInterface").fill(
            "#LegalNativePlace", Agent["LegalNativePlace"])
        self.page13.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", Agent["LegalNativePlace"])
        self.page13.wait_for_timeout(2000)
        self.page13.frame(name="fraInterface").press(
            "#LegalNativePlace", "Enter")

        # 保存
        self.page13.frame(name="fraInterface").click("#save")
        # 返回
        self.page13.frame(name="fraInterface").click("#return")

# 審閱期說明書(NTWL0202)
    def Risk(self, SaleReport):
        # time.sleep(2)
        self.page2.touchscreen.tap(100, 100)
        self.page2.once("dialog", lambda dialog: dialog.accept())
        with self.page2.expect_popup() as popup_info:
            self.page2.frame(name="fraInterface").click("#riskbutton57")
        self.page15 = popup_info.value
        self.page15.wait_for_load_state()
        print(self.page15.title())
        # self.page15.screenshot(path="審閱期間說明書.png")

        # 選擇審閱聲明
        self.page15.frame(name="fraInterface").dblclick("#ReviewStatement")
        # self.page15.frame(name="fraInterface").fill("#ReviewStatement", SaleReport["ReviewStatement"])
        self.page15.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", SaleReport["ReviewStatement"])
        self.page15.frame(name="fraInterface").press(
            "#ReviewStatement", "Enter")

        # 條款提供日期
        self.page15.frame(name="fraInterface").dblclick("#ProvisionDate")
        self.page15.frame(name="fraInterface").fill(
            "#ProvisionDate", SaleReport["ProvisionDate"])

        # 選擇聲明日期
        self.page15.frame(name="fraInterface").dblclick("#StatementDate")
        self.page15.frame(name="fraInterface").fill(
            "#StatementDate", SaleReport["StatementDate"])

        # 選擇險種 ProductListGrid1r0
        self.page15.frame(name="fraInterface").dblclick("#ProductListGrid1r0")
        self.page15.frame(name="fraInterface").press(
            "#ProductListGrid1r0", "Enter")

        # 保存
        self.page15.frame(name="fraInterface").click("input[name=\"sure\"]")

        # self.page15.screenshot(path="審閱期間說明書.png")

        self.page15.wait_for_timeout(1000)

        # 返回
        self.page15.frame(name="fraInterface").click("text=返回")

        self.page15.wait_for_timeout(1000)

# 以外幣收付之非投資型人身保險客戶適合度調查評估表
    def USD_report(self):
        # time.sleep(2)
        self.page2.touchscreen.tap(100, 100)
        self.page2.once("dialog", lambda dialog: dialog.accept())
        with self.page2.expect_popup() as popup_info:
            self.page2.frame(name="fraInterface").click("#riskbutton55")
        self.page16 = popup_info.value
        self.page16.wait_for_load_state()
        self.page16.screenshot(path="USDreport.png")
        print(self.page16.title())

        self.page16.frame(name="fraInterface").dblclick(
            "#Investigation1Grid6r0")
        # self.page16.frame(name="fraInterface").fill("#Investigation1Grid6r0", "Y")
        self.page16.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", "Y")
        self.page16.frame(name="fraInterface").press(
            "#Investigation1Grid6r0", "Enter")

        self.page16.frame(name="fraInterface").dblclick("#YesOrNo")
        self.page16.frame(name="fraInterface").fill("#YesOrNo", "Y")
        self.page16.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", "Y")
        self.page16.frame(name="fraInterface").press("#YesOrNo", "Enter")

        self.page16.frame(name="fraInterface").dblclick(
            "#Investigation2Grid5r0")
        self.page16.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", "Y")
        self.page16.frame(name="fraInterface").press(
            "#Investigation2Grid5r0", "Enter")

        self.page16.frame(name="fraInterface").dblclick(
            "#Investigation2Grid5r1")
        self.page16.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", "Y")
        self.page16.frame(name="fraInterface").press(
            "#Investigation2Grid5r1", "Enter")

        self.page16.frame(name="fraInterface").dblclick(
            "#Investigation2Grid5r2")
        self.page16.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", "Y")
        self.page16.frame(name="fraInterface").press(
            "#Investigation2Grid5r2", "Enter")

        self.page16.frame(name="fraInterface").dblclick(
            "#Investigation2Grid5r3")
        self.page16.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", "Y")
        self.page16.frame(name="fraInterface").press(
            "#Investigation2Grid5r3", "Enter")

        self.page16.frame(name="fraInterface").dblclick(
            "#Investigation2Grid5r4")
        self.page16.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", "Y")
        self.page16.frame(name="fraInterface").press(
            "#Investigation2Grid5r4", "Enter")

    # 保存
        self.page16.frame(name="fraInterface").click(
            "#InvestigationSaveButton")
    # 關閉
        self.page16.frame(name="fraInterface").click("#closeSaveButton")

# 銀行轉帳授權書
    def Authorization_Letter(self, personinfo, Salesman_value):
        # time.sleep(2)
        self.page2.touchscreen.tap(100, 100)
        self.page2.once("dialog", lambda dialog: dialog.accept())
        with self.page2.expect_popup() as popup_info:
            self.page2.frame(name="fraInterface").click("#BankLoad2")
        self.page17 = popup_info.value
        self.page17.wait_for_load_state()
        self.page17.screenshot(path="USDreport.png")
        print(self.page17.title())

        # # 得到保單號碼 #Renewaln2Grid1r0
        value1 = self.page17.frame(name="fraInterface").wait_for_selector(
            "//html/body/form/table[2]/tbody/tr/td/span/div[1]/div[5]/table/tbody/tr/td[3]/div/input")
        # value=value1.get_attribute('name')
        Policy_value = value1.input_value()
        print("我是保單號碼2:"+Policy_value)

        # 填入授權書代碼
        self.page17.frame(name="fraInterface").fill(
            "#AuthorBar", "ML"+Policy_value)

        # 輸入申請日期 BankApplyDate
        self.page17.frame(name="fraInterface").fill(
            "#AppDate", personinfo["BankApplyDate"])
        # 選擇授權書類型 X續期保費 SX首續期保費
        self.page17.frame(name="fraInterface").dblclick("#AuthorType")
        self.page17.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", personinfo["AuthorType"])
        self.page17.frame(name="fraInterface").press("#AuthorType", "Enter")

        # 填入業務員代碼
        self.page17.frame(name="fraInterface").fill(
            "#AgentCode", Salesman_value)

        # 授權人與保單關係
        if(personinfo["Renewaln2Grid5r0"] == '1'):
            self.page17.frame(name="fraInterface").dblclick(
                "#Renewaln2Grid5r0")
            self.page17.frame(name="fraInterface").select_option(
                "select[name=\"codeselect\"]", "1")
            self.page17.frame(name="fraInterface").click(
                "//html/body/span/select/option[1]")
        elif(personinfo["Renewaln2Grid5r0"] == '10'):
            self.page17.frame(name="fraInterface").dblclick(
                "#Renewaln2Grid5r0")
            self.page17.frame(name="fraInterface").select_option(
                "select[name=\"codeselect\"]", "10")
            self.page17.frame(name="fraInterface").click(
                "//html/body/span/select/option[2]")

        # 是否檢附關係説明文件
        self.page17.frame(name="fraInterface").dblclick("#Renewaln2Grid6r0")
        self.page17.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", "1")
        self.page17.frame(name="fraInterface").click(
            "//html/body/span/select/option[1]")

        # 點擊檢核
        self.page17.frame(name="fraInterface").click("text=檢核")

        # 點擊保存
        self.page17.frame(name="fraInterface").click("text=保存")

        # 選擇幣別
        self.page17.frame(name="fraInterface").dblclick("#AuthorCurrency")
        self.page17.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", personinfo["AuthorCurrency"])
        self.page17.frame(name="fraInterface").press(
            "#AuthorCurrency", "Enter")

        # 選擇銀行授權代碼 8150015日盛銀行  0095185彰銀總部(彰化銀行)
        self.page17.frame(name="fraInterface").dblclick("#BankCode")
        self.page17.frame(name="fraInterface").select_option(
            "select[name=\"codeselect\"]", personinfo['bankcode'])
        self.page17.frame(name="fraInterface").press("#BankCode", "Enter")
        # if(personinfo['bankcode']=='8150015'):
        #     self.page17.frame(name="fraInterface").dblclick("#BankCode")
        #     self.page17.frame(name="fraInterface").select_option("select[name=\"codeselect\"]", "8150015")
        #     self.page17.frame(name="fraInterface").press("#BankCode", "Enter")
        # elif(personinfo['bankcode']=='0095185'):
        #     self.page17.frame(name="fraInterface").dblclick("#BankCode")
        #     self.page17.frame(name="fraInterface").select_option("select[name=\"codeselect\"]", "0095185")
        #     self.page17.frame(name="fraInterface").press("#BankCode", "Enter")
        # elif(personinfo['bankcode']=='1080014'):
        #     self.page17.frame(name="fraInterface").dblclick("#BankCode")
        #     self.page17.frame(name="fraInterface").select_option("select[name=\"codeselect\"]", "1080014")
        #     self.page17.frame(name="fraInterface").press("#BankCode", "Enter")
        # # 玉山銀行
        # elif(personinfo['bankcode']=='8080015'):
        #     self.page17.frame(name="fraInterface").dblclick("#BankCode")
        #     self.page17.frame(name="fraInterface").select_option("select[name=\"codeselect\"]", "8080015")
        #     self.page17.frame(name="fraInterface").press("#BankCode", "Enter")

        # 如果關係選"其他"則進來輸入授權人姓名以及ID
        if(personinfo["Renewaln2Grid5r0"] == '10'):
            # 授權人姓名
            self.page17.frame(name="fraInterface").fill(
                "#AuthorName", personinfo["AppntName"]+"用")
            # 證件類型
            self.page17.frame(name="fraInterface").dblclick("#AppntIDType")
            self.page17.frame(name="fraInterface").select_option(
                "select[name=\"codeselect\"]", "0")
            self.page17.frame(name="fraInterface").press(
                "#AppntIDType", "Enter")
            # # 授權人生日
            self.page17.frame(name="fraInterface").fill(
                "#AuthorBirth", personinfo["AppntBirthday"])
            # # 授權人ID(兩次)
            self.page17.frame(name="fraInterface").fill(
                "#AuthorID", personinfo["AuthorID"])
            # 授權人姓名
            self.page17.frame(name="fraInterface").fill(
                "#AuthorName", personinfo["AppntName"]+"用")
            # self.page17.frame(name="fraInterface").press("#AuthorID", "Enter")
            self.page17.frame(name="fraInterface").fill(
                "#AuthorID", personinfo["AuthorID"])

        # 輸入帳號(兩次)
        self.page17.frame(name="fraInterface").fill(
            "#AuthorAccNo", personinfo["AuthorAccNo"])
        # 點擊授權人電話
        self.page17.frame(name="fraInterface").click("#AuthorPhone")
        # self.page17.frame(name="fraInterface").press("#AuthorAccNo", "Enter")
        self.page17.frame(name="fraInterface").fill(
            "#AuthorAccNo", personinfo["AuthorAccNo"])

        # 檢核
        self.page17.frame(name="fraInterface").dispatch_event(
            "//html/body/form/a[4]", "click")
        # self.page17.wait_for_timeout(100000)

        # 點擊應付未付
        self.page17.frame(name="fraInterface").dispatch_event(
            "//html/body/form/a[5]", "click")

        # 保存
        self.page17.frame(name="fraInterface").dispatch_event(
            "//html/body/form/a[6]", "click")
        self.page17.wait_for_timeout(1000)
        self.page17.close()
