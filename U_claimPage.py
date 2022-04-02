from time import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

# 理賠
class ClaimPage():
    ### 新增受理按鈕
    newApply=(By.PARTIAL_LINK_TEXT,'新增受理')

    ### 輸入身分證字號 
    IdNo=(By.ID,'IdNo')

    ### 選擇送件來源 
    RgtSourcesType=(By.ID,'RgtSourcesType')
    codeselect=(By.ID,'codeselect')

    ### 收件日期
    AcceptedDate=(By.ID,'AcceptedDate')

    ### 事故日期
    AccidentDate=(By.ID,'AccidentDate')

    ### 出險原因
    occurReason=(By.ID,'occurReason')

    # 診斷病名(跳轉頁面)
    AccReason=(By.ID,'AccReason')
    LLColDiseaseDetailGrid1r0=(By.ID,'LLColDiseaseDetailGrid1r0')

    # 存檔並返回
    saveBtn=(By.XPATH,'//html/body/form/table[2]/tbody/tr/td/input')

    


    def click_ClaimPage_newApplyBtn(self):
        self.find_element(*self.newApply).click()

    def input_ClaimPage_IdNo(self,personinfo):
        self.find_element(*self.IdNo).send_keys(personinfo['AppntIDNo'])

    def select_ClaimPage_RgtSourcesType(self):
        self.find_element(*self.RgtSourcesType).click()
        b=self.find_element(*self.codeselect)
        ### 0業代 1保經代 2郵寄 3櫃檯 4I-BON 5區塊鏈 6龍e賠 7醫院扣抵
        Select(b).select_by_value('0')
        # self.find_element(*self.RgtSourcesType).send_keys(Keys.ENTER)

    def input_ClaimPage_AcceptedDate(self):
        self.find_element(*self.AcceptedDate).send_keys('2021-11-29')

    def input_ClaimPage_AccidentDate(self):
        self.find_element(*self.AccidentDate).send_keys('2021-11-29')
    
    def select_ClaimPage_occurReason(self):
        self.find_element(*self.occurReason).click()
        Reason=self.find_element(*self.codeselect)
        ### 1-意外 2-疾病 3-自殺
        Select(Reason).select_by_value('3')

    def click_ClaimPage_AccReason(self):
        # t0=self.driver.title
        # print("目前窗口："+t0)
        self.find_element(*self.AccReason).click()

        self.switch_window(1)
        t1=self.driver.title
        print("目前窗口："+t1)

        # ### 輸入診斷病名代碼 
        self.find_element(*self.LLColDiseaseDetailGrid1r0).send_keys('Z999')
        AccReason=self.find_element(*self.codeselect)
        ### Z999-自殺
        Select(AccReason).select_by_value('Z999')
        self.find_element(*self.LLColDiseaseDetailGrid1r0).send_keys(Keys.ENTER)

        # 存檔並返回
        # saveBtn1 = self.find_element(*self.saveBtn)
        # print("我是"+saveBtn1[0])
        # saveBtn2 = saveBtn1[1]
        # ActionChains(self.driver).move_to_element(saveBtn2).click().perform()
        # test=self.find_element(*self.saveBtn).get_attribute('value')
        # print(test)

        self.find_element(*self.saveBtn).click()