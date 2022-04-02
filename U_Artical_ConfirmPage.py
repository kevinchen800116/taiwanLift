from time import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains


# 承保處理-->個人保單-->人工核保頁面
class UArtical_ConfirmPage():
    # 第一個査詢書入保單號碼
    PublicWorkPoolQueryGrid1r0=(By.ID,'PublicWorkPoolQueryGrid1r0')

    # 查詢按鈕
    publicSearch=(By.ID,'publicSearch')

    # 水池內按鈕 點了會跳轉業面
    PublicWorkPoolGridSel0=(By.ID,'PublicWorkPoolGridSel0')

# ------------------------------------------------
    # 核保提示信息
    # 全選
    # checkAllUWErrGrid=(By.ID,'checkAllUWErrGrid')
    checkAllUWErrGrid=(By.CSS_SELECTOR,'.mulinetitle')

    # 五個
    UWErrGridChk0=(By.ID,'UWErrGridChk0')
    UWErrGridChk1=(By.ID,'UWErrGridChk1')
    UWErrGridChk2=(By.ID,'UWErrGridChk2')
    UWErrGridChk3=(By.ID,'UWErrGridChk3')
    UWErrGridChk4=(By.ID,'UWErrGridChk4')

    # 解除按鈕
    AutoUWButton=(By.ID,'AutoUWButton')

    # 核保結論(選擇框)12345
    auwState=(By.ID,'auwState')

    # 核保流向(選擇框)014
    uwUpReport=(By.ID,'uwUpReport')

    # 核保結論&核保流向 (共用選擇框)
    select_value=(By.ID,'codeselect')

    # 核保意見(input框)
    UWIdea=(By.ID,'UWIdea')

    # 人工核保 確定按鈕
    button1=(By.ID,'button1')

# ---------------以下個變數還未做--------------------------------
    
    # 第二個査詢 輸入保單號碼
    PrivateWorkPoolQueryGrid2r0=(By.ID,'PrivateWorkPoolQueryGrid2r0')

    # 查詢按鈕
    privateSearch=(By.ID,'privateSearch')

    # 水池內按鈕 點了會跳轉業面
    PrivateWorkPoolGridSel0=(By.ID,'PrivateWorkPoolGridSel0')

    # 選擇核保結論方法
    def sele_ArtiConfirmPage_auwState(self):
        self.find_element(*self.auwState).click()
        s=self.find_element(*self.select_value)
        # 固定選擇1
        Select(s).select_by_value('1')

    #選擇核保流向方法 
    def sele_ArtiConfirmPage_uwUpReport(self):
        self.find_element(*self.uwUpReport).click()
        s=self.find_element(*self.select_value)
        # 固定選擇0
        Select(s).select_by_value('0')

    # 第一個査詢書入保單號碼
    def input_ArtiConfirmPage_first_PolicyNumber(self,Policy_number):
        self.find_element(*self.PublicWorkPoolQueryGrid1r0).send_keys(Policy_number)

    # 第一個查詢按鈕
    def click_ArtiConfirmPage_first_QBtn(self):
        self.find_element(*self.publicSearch).click()

    # 公用池選擇
    def click_ArtiConfirmPage_PublicPool(self):
        self.find_element(*self.PublicWorkPoolGridSel0).click()

    # 核保訊息全選
    def click_ArtiConfirmPage_checkAll(self):
        # 跳轉視窗
        self.switch_window(1)
        # 確定跳轉視窗的名稱
        print(self.driver.title)
        # 轉換網頁內容
        self.switch_default()
        self.switch_frame("fraInterface")
        # 點擊全選
        self.find_element(*self.checkAllUWErrGrid).click()

    # 核保訊息(點選核保訊息第1-5個)
    def click_ArtiConfirmPage_check1to5(self):
        self.switch_window(1)
        print(self.driver.title)
        self.switch_default()
        self.switch_frame("fraInterface")
        self.find_element(*self.UWErrGridChk0).click()
        self.find_element(*self.UWErrGridChk1).click()
        self.find_element(*self.UWErrGridChk2).click()
        self.find_element(*self.UWErrGridChk3).click()
        self.find_element(*self.UWErrGridChk4).click()

    # 解除按鈕
    def click_ArtiConfirmPage_AutoUWButton(self):
        self.find_element(*self.AutoUWButton).click()

    # 輸入核保意見
    def input_ArtiConfirmPage_UWIdea(self):
        # 固定輸入
        # self.find_element(*self.AutoUWButton).click()
        self.find_element(*self.UWIdea).send_keys('DD')

    # 點擊確定
    def click_ArtiConfirmPage_button1(self):
        self.find_element(*self.button1).click()

    # 第二個輸入保單號碼 PrivateWorkPoolQueryGrid2r0
    def input_ArtiConfirmPage_second_PoliceNumber(self,Policy_number):
        self.find_element(*self.PrivateWorkPoolQueryGrid2r0).click()
        self.find_element(*self.PrivateWorkPoolQueryGrid2r0).send_keys(Policy_number)

    # 第二個查詢
    def click_ArtiConfirmPage_second_Qbtn(self):
        self.find_element(*self.privateSearch).click()

    # 第二個水池按鈕
    def click_ArtiConfirmPage_second_PrivatePool(self):
        self.find_element(*self.PrivateWorkPoolGridSel0).click()