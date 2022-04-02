from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import time

# 保費-->單筆自繳-->單筆自繳建檔頁面
class PremiumPage():

    # 共用選擇框
    select_value=(By.ID,'codeselect')

    # 單據類型
    DocumentType2=(By.ID,'DocumentType2')

    # 立帳流水號
    AccountNo=(By.ID,'AccountNo')
    # 新增按鈕
    PremiumPage_AddBtn=(By.PARTIAL_LINK_TEXT,'新建')
  #----------------------------------------------------
    # 單據金額(現金 1)
    DocumentMoney=(By.ID,'DocumentMoney')
    # 單據金額(匯款 3)
    DocumentMoney3=(By.ID,'DocumentMoney3')
    # 使用金額
    UseMoney=(By.ID,'UseMoney')
    # 繳款日期(現金)
    EnteraccDate=(By.ID,'EnteraccDate')
    # 繳款日期(匯款)
    EnteraccDate3=(By.ID,'EnteraccDate3')
    # 保存按鈕
    savebtn=(By.PARTIAL_LINK_TEXT,'保存')
    # 單據訊息結果
    PDFGridSel0=(By.ID,'PDFGridSel0')
    # 繳費類型
    PayType=(By.ID,'PayType')
    # 首期保費
    BusinessType=(By.ID,'BusinessType')
    # 保單號碼
    ContNoOrAppntID=(By.ID,'ContNoOrAppntID')

    # 繳款人與保單關係
    Relation=(By.ID,'Relation')
    # 同要保人
    IsAppnt=(By.ID,'IsAppnt')
    # 分配金額(現金)
    PDF2Grid8r0=(By.ID,'PDF2Grid8r0')

    # 分配金額(匯款) PDF2Grid10r0
    # /html/body/form/div[7]/table/tbody/tr/td/span/div/div[5]/table/tbody/tr/td[12]/div/input
    # PDF2Grid10r0=(By.ID,'PDF2Grid10r0')
    PDF2Grid10r0=(By.XPATH,'/html/body/form/div[7]/table/tbody/tr/td/span/div/div[5]/table/tbody/tr/td[12]/div/input')

    # 全部提交
    savebtn3=(By.PARTIAL_LINK_TEXT,'全部提交')

    # 銀行代號
    BankCode=(By.ID,'BankCode3')
    # 銀行帳號
    BankAcc=(By.ID,'BankAcc3')
    # 幣別
    Currency=(By.ID,'Currency')
    # 繳款人帳號
    PayAcc=(By.ID,'PayAcc')


    # 選擇單據類型
    def select_PremiumPage_DocumentType_cash(self,Acc_sub):
        if(Acc_sub=='1001002'):
            # 點擊單據類型叫出選擇框
            self.find_element(*self.DocumentType2).click()
            # 選擇現金(01) 、02劃撥 、03匯款 、18支票
            b=self.find_element(*self.select_value)
            Select(b).select_by_value('01')
        elif(Acc_sub=='149000106'):
            # 點擊單據類型叫出選擇框
            self.find_element(*self.DocumentType2).click()
            # 選擇現金(01) 、02劃撥 、03匯款 、18支票
            b=self.find_element(*self.select_value)
            Select(b).select_by_value('03')

    # 新建按鈕
    def click_PremiumPage_AddBtn(self):
        self.find_element(*self.PremiumPage_AddBtn).click()

  # ----單據建檔--------------------------------------------------------------

    # 單據金額(現金)
    def input_PremiumPage_DocumentMoney(self,price):
        self.find_element(*self.DocumentMoney).send_keys(price)
    # 單據金額(匯款)
    def input_PremiumPage_DocumentMoney3(self,price):
        self.find_element(*self.DocumentMoney3).send_keys(price)

    # 使用金額(目前可能被移除)
    def input_PremiumPage_UseMoney(self,price):
        self.find_element(*self.UseMoney).send_keys(price)

    # 繳款日期(現金)
    def input_PremiumPage_EnteraccDate(self,personinfo):
        self.find_element(*self.EnteraccDate).send_keys(personinfo["PolAppntDate"])

    # 繳款日期(匯款)
    def input_PremiumPage_EnteraccDate3(self,personinfo):
        self.find_element(*self.EnteraccDate3).send_keys(personinfo["PolAppntDate"])

    # 輸入立帳流水號
    def input_PremiumPage_AccountNo(self,ASN):
        self.find_element(*self.AccountNo).send_keys(ASN)

    # 輸入銀行代號(花旗外幣)
    def input_PremiumPage_BankCode(self):
        self.find_element(*self.BankCode).click()
        self.find_element(*self.BankCode).send_keys('0210018')
        # 選擇銀行代號
        b=self.find_element(*self.select_value)
        Select(b).select_by_value('0210018')
        testdata=self.find_element(*self.select_value)
        ActionChains(self.driver).move_to_element(testdata).click().perform()

    # 輸入銀行帳號 & 輸入繳款人帳號
    def input_PremiumPage_BankAcc(self):
        self.find_element(*self.BankAcc).click()
        self.find_element(*self.BankAcc).send_keys('5446943202')
        # 選擇銀行代號
        b=self.find_element(*self.select_value)
        # 5446943202
        Select(b).select_by_value('2')
        testdata=self.find_element(*self.select_value)
        ActionChains(self.driver).move_to_element(testdata).click().perform()
        # 選擇幣別
        self.find_element(*self.Currency).click()
        b=self.find_element(*self.select_value)
        Select(b).select_by_value('USD')
        # testdata2=self.find_element(*self.select_value)
        # ActionChains(self.driver).move_to_element(testdata2).click().perform()
        # 輸入繳款人帳號
        self.find_element(*self.PayAcc).send_keys('123456789012')

    # 保存
    def click_PremiumPage_savebtn(self):
        self.find_element(*self.savebtn).click()

    # 單據訊息結果
    def click_PremiumPage_SelectResult(self):
        self.find_element(*self.PDFGridSel0).click()

    # 繳費類型
    def select_PremiumPage_PayType(self):
        # 點擊繳費類型
        self.find_element(*self.PayType).click()
        # 選擇1繳費  2餘額帳戶
        b=self.find_element(*self.select_value)
        Select(b).select_by_value('1')
        self.find_element(*self.PayType).send_keys(Keys.ENTER)


    # 首期保費(現金)
    def select_PremiumPage_BusinessType(self,price,Policy_number):
        # 點擊業務類型
        self.find_element(*self.BusinessType).click()
        # 選擇1繳費  2餘額帳戶
        b=self.find_element(*self.select_value)
        Select(b).select_by_value('01')

        # 保單號碼
        self.find_element(*self.ContNoOrAppntID).send_keys(Policy_number)

        # 同要保人
        self.find_element(*self.IsAppnt).click()

        # 保存按鈕
        btn=self.find_elements(*self.savebtn)
        btn1=btn[1]
        ActionChains(self.driver).move_to_element(btn1).click().perform()

        # 分配金額
        self.find_element(*self.PDF2Grid8r0).click()
        self.find_element(*self.PDF2Grid8r0).send_keys(price)
        # 全部提交
        self.find_element(*self.savebtn3).click()

        # # 轉換視窗
        # self.switch_to.window(handles[0])

    # 首期保費(匯款)
    def select_PremiumPage_BusinessType3(self,price,Policy_number):
        # 點擊業務類型
        self.find_element(*self.BusinessType).click()
        # 選擇1繳費  2餘額帳戶
        b=self.find_element(*self.select_value)
        Select(b).select_by_value('01')

        # 保單號碼
        self.find_element(*self.ContNoOrAppntID).send_keys(Policy_number)

        # 繳款人與保單關係
        Relation2=self.find_element(*self.Relation)
        ActionChains(self.driver).move_to_element(Relation2).double_click().perform()
        # 選擇1要保人
        x=self.find_element(*self.select_value)
        Select(x).select_by_value('1')
        # self.find_element(*self.x).send_keys(Keys.ENTER)


        # 保存按鈕
        btn=self.find_elements(*self.savebtn)
        btn1=btn[1]
        ActionChains(self.driver).move_to_element(btn1).click().perform()

        # alert被拿掉了==
        # self.switch_to_alert_accept()

        # 分配金額 PDF2Grid10r0 
        self.find_element(*self.PDF2Grid10r0).click()
        self.find_element(*self.PDF2Grid10r0).send_keys(price)
        # 全部提交
        self.find_element(*self.savebtn3).click()
        

# 保費-->單筆自繳-->單筆自繳審核頁面
class PremiumCheckPage():

    # 金額範圍上限
    Amount=(By.ID,'Amount')
    # 金額範圍下限
    Amount2=(By.ID,'Amount2')
    # 幣別
    Currency=(By.ID,'Currency')
    # 繳費方式
    DocumentType=(By.ID,'DocumentType')
    # 銀行入帳起日
    StartPayDate=(By.ID,'StartPayDate')
    # 銀行入帳迄日
    EndPayDate=(By.ID,'EndPayDate')
    # 查詢按鈕
    PremiumCheckQbtn=(By.CSS_SELECTOR,'a.button')
    # pool first one
    TARReviewGridChk0=(By.ID,'TARReviewGridChk0')

    # 保存成功訊息
    contentTD=(By.ID,'contentTD')

    # 金額範圍上限
    def input_PremiumCheckPage_Amount(self,price):
        self.find_element(*self.Amount).send_keys(price)
    # 金額範圍下限
    def input_PremiumCheckPage_Amount2(self,price):
        self.find_element(*self.Amount2).send_keys(price)

    def select_PremiumCheckPage_Currency(self):
        # 點擊幣別
        self.find_element(*self.Currency).click()
        time.sleep(1)
        # 選擇台幣
        b=self.find_element(*self.select_value)
        Select(b).select_by_value('NTD')
        # self.find_element(*self.Currency).send_keys(Keys.ENTER)

    def select_PremiumCheckPage_DocumentType(self):
        # 點擊繳費方式DocumentType
        self.find_element(*self.DocumentType).click()
        # 選擇01現金  02劃撥  03匯款
        b=self.find_element(*self.select_value)
        Select(b).select_by_value('01')
        # self.find_element(*self.Currency).send_keys(Keys.ENTER)

    # 銀行入帳起日
    def input_PremiumCheckPage_StartPayDate(self,personinfo):
        self.find_element(*self.StartPayDate).send_keys(personinfo["PolAppntDate"])

    # 銀行入帳迄日
    def input_PremiumCheckPage_EndPayDate(self,personinfo):
        self.find_element(*self.EndPayDate).send_keys(personinfo["PolAppntDate"])

    def click_PremiumCheckPage_Qbtn(self):
        q=self.find_elements(*self.PremiumCheckQbtn)
        q1=q[0]
        ActionChains(self.driver).move_to_element(q1).click().perform()

    def click_PremiumCheckPage_PoolFirst(self):
        self.find_element(*self.TARReviewGridChk0).click()

    def submit_PremiumCheckPage_All(self):
        q=self.find_elements(*self.PremiumCheckQbtn)
        q1=q[2]
        ActionChains(self.driver).move_to_element(q1).click().perform()
        # # 以下為練習assert
        self.switch_window(1)
        success=self.find_element(*self.contentTD).text
        return success
        # assert success == "操作成功。"

    