import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select

# 出納-->立帳作業-->現金/匯款/劃撥立帳-->單筆立帳頁面&立帳綜合查詢
class AccountPage():
    bank_rc_date_btn=(By.ID,'EnterAccDate')# 銀行收款日期
    Money=(By.ID,'Money')  # 金額 or 立帳流水號查詢頁面的開始金額
    Remark=(By.ID,'Remark')# 摘要(需放入保單號碼)
    select_AccountCode=(By.ID,'AccountCode')# 會計科目
    select_AccountCode_value=(By.ID,'codeselect')# 會計科目
    add_btn=(By.LINK_TEXT,'新增')# 新增按鈕
    setup_btn=(By.LINK_TEXT,'立帳')# 立帳按鈕

# -----------立帳流水號查詢頁面--------------------------------------------
    Currency=(By.ID,'Currency')# 幣別
    toMoney=(By.ID,'toMoney')# 最後的金額

    # 立帳狀態
    State=(By.ID,'State')
    State_codeselect=(By.ID,'codeselect')

    # querybtn=(By.PARTIAL_LINK_TEXT,'查詢')# 查詢按鈕
    querybtn=(By.CSS_SELECTOR,'a.button')
    AccSerNumber=(By.ID,'RBResultGrid1r0')# 立帳流水號第一欄的值

    # def click_bank_re_date(self):
    #     self.find_element(*self.bank_rc_date_btn).click()

# --------單筆立帳方法------------------------------------------------------
    def input_Account_Money(self,price):
        self.find_element(*self.Money).send_keys(price)

    def input_Account_Policy(self,Policy_number):
        self.find_element(*self.Remark).send_keys(Policy_number)

    def select_Account_subject(self,Acc_sub):
        self.find_element(*self.select_AccountCode).click()
        s=self.find_element(*self.select_AccountCode_value)
        Select(s).select_by_value(Acc_sub)

    def click_Account_add_btn(self):
        self.find_element(*self.add_btn).click()

    def click_Account_setup_btn(self):
        self.find_element(*self.setup_btn).click()

# --------------立帳流水號查詢方法----------------------------------------
    def query_AccountPage_serial_number(self,price,Acc_sub):
        # 點擊幣別叫出選擇框
        self.find_element(*self.Currency).click()

        if(Acc_sub=='149000106'):
            # 選擇USD
            b=self.find_element(*self.select_AccountCode_value)
            # Select(b).select_by_value(CurrencyCode)
            Select(b).select_by_value('USD')
            print(Acc_sub+"USD選擇成功")
        elif(Acc_sub=='1001002'):
            # 選擇台幣
            b=self.find_element(*self.select_AccountCode_value)
            # Select(b).select_by_value(CurrencyCode)
            Select(b).select_by_value('NTD')
            print(Acc_sub+"台幣選擇成功")

        # 輸入起始金額
        self.find_element(*self.Money).send_keys(price)
        # 輸入結束金額
        self.find_element(*self.toMoney).send_keys(price)

        # 立帳狀態
        self.find_element(*self.State).click()
        time.sleep(1)
        # 選擇"未沖銷"
        h=self.find_element(*self.State_codeselect)
        Select(h).select_by_value('03')

        # 點擊查詢按鈕
        self.find_element(*self.querybtn).click()

        # 回傳第一欄的立帳流水號碼
        Account_serial_number=self.find_element(*self.AccSerNumber).get_attribute('value')
        return Account_serial_number

