from selenium.webdriver.common.by import By
from UAT_BasePage import BasePage
from U_AutoConfirmPage import UAutoConfirmPage
from U_QueryPage import QueryPage
from U_AccountPage import AccountPage
from U_PremiumPage import PremiumPage,PremiumCheckPage
from U_Artical_ConfirmPage import UArtical_ConfirmPage
from U_SignPage import USignPage, UPrintSignPage, USignBackPage, TaskPage
from U_PreservationPage import PreservationPage, Preservation_rejectPage
from U_claimPage import ClaimPage
from U_Renew import Renew, Bank_Authorization, Bank_Delivery, Bank_Approval, Bank_Return
from NewInput import NewInput



#继承BasePage类
class LoginPage(BasePage, UAutoConfirmPage, QueryPage, AccountPage, PremiumPage, PremiumCheckPage, 
UArtical_ConfirmPage, USignPage, PreservationPage, UPrintSignPage, ClaimPage, USignBackPage, TaskPage, Renew, Bank_Authorization, Bank_Approval, Bank_Delivery, Bank_Return,
Preservation_rejectPage, NewInput):
    #定位器，通过元素属性定位元素对象
    username_loc =(By.CSS_SELECTOR,'input#UserCode2')
    password_loc =(By.CSS_SELECTOR,'input#PWD2')
    submit_loc =(By.ID,'submit2')
    # span_loc =(By.CSS_SELECTOR,"div.error-tt>p")
    # dynpw_loc =(By.ID,"lbDynPw")
    # userid_loc =(By.ID,"spnUid")
    
    #操作
    #通过继承覆盖（Overriding）方法：如果子类和父类的方法名相同，优先用子类自己的方法。
    #打开网页
    def open(self):
    #调用page中的_open打开连接
        self._open(self.base_url, self.pagetitle)
    
    #输入用户名：调用send_keys对象，输入用户名
    def input_username(self, username):
#        self.find_element(*self.username_loc).clear()
        print('進來了'+username)
        self.find_element(*self.username_loc).send_keys(username)
    
    #输入密码：调用send_keys对象，输入密码
    def input_password(self, password):
#        self.find_element(*self.password_loc).clear()
        self.find_element(*self.password_loc).send_keys(password)
        
    #点击登录：调用send_keys对象，点击登录
    def click_submit(self):
        self.find_element(*self.submit_loc).click()
    
    # #用户名或密码不合理是Tip框内容展示
    # def show_span(self):
    #     return self.find_element(*self.span_loc).text
    
    # #切换登录模式为动态密码登录（IE下有效）
    # def swich_DynPw(self):
    #     self.find_element(*self.dynpw_loc).click()
    
    # #登录成功页面中的用户ID查找
    # def show_userid(self):
    #     return self.find_element(*self.userid_loc).text