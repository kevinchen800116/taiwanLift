from selenium.webdriver.common.by import By


# 承保處理-->個人保單-->自動核保頁面
class UAutoConfirmPage():
    # 保單號碼輸入框
    PolicyNb=(By.ID,'PublicWorkPoolQueryGrid1r0')
    # 查詢按鈕
    check=(By.ID,'publicSearch')
    # 選擇公共工作池內的第一個保單號碼的圈圈
    select=(By.ID,'PublicWorkPoolGridChk0')
    # 自動核保按鈕
    autoSubmit=(By.CSS_SELECTOR,'input.cssButton')


    # 保單號碼 Policy_number
    def input_Policy_number(self,Policy_number):
        self.find_element(*self.PolicyNb).send_keys(Policy_number)

    # 查詢 publicSearch
    def click_Auto_Confirm_Search(self):
        self.find_element(*self.check).click()

	# 公共工作池第一個圈圈
    def click_Auto_Confirm_select(self):
        self.find_element(*self.select).click()
        
	# 自動核保按鈕
    def click_Auto_Confirm_Submit(self):
        self.find_element(*self.autoSubmit).click()

    

