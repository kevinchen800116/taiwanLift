from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from BasePage import BasePage

class AddPage(BasePage):
    P=(By.PARTIAL_LINK_TEXT,'承保處理')
    InsideMenu=(By.CSS_SELECTOR,'.goon')
    R=(By.PARTIAL_LINK_TEXT,'新單覆核')
    
    def Insurance_handle(self):
        self.find_element(*self.P).click()

    def Insurance_Inside_Menu(self):
        goon = self.find_elements(*self.InsideMenu)
        goon1 = goon[0]
        ActionChains(self.driver).move_to_element(goon1).click().perform()
        self.find_element(*self.R).click()
