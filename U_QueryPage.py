from selenium.webdriver import support
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import time
import datetime


# 查詢頁面(新單查詢&新單覆核)
class QueryPage():
    # # 成功訊息
    # contentTD=(By.ID,'contentTD')

    # 左側menu共用區塊
    Menu=(By.CSS_SELECTOR,'.goon')
    # 往旁邊滑的按鈕
    menuButton_next=(By.CSS_SELECTOR,'li.next')
    menuButton_prev=(By.CSS_SELECTOR,'li.prev')

    menuButton =(By.PARTIAL_LINK_TEXT,'承保處理')

    # menuButton2=(By.CSS_SELECTOR,'li.show-poster-3')
    menuButton2=(By.PARTIAL_LINK_TEXT,'綜合列印')

    menuButton3=(By.PARTIAL_LINK_TEXT,'保全處理')

    menuButton7=(By.PARTIAL_LINK_TEXT,'綜合管理')

    # 保全案件調撥
    preser_return=(By.CSS_SELECTOR,'a.innera')

    # 契變崗位
    ChangePost=(By.ID,'ChangePost')
    # 選擇框
    codeselect=(By.ID,'codeselect')
    # 個人池選項 0000000019
    checkAllPersonalWorkPoolGrid=(By.ID,'PersonalWorkPoolGridChk0')
    # 查詢按鈕
    preser_return_Btn=(By.CSS_SELECTOR,'a.button')
    # 處理人員
    ProcessOperater=(By.ID,'ProcessOperater')

    # 無掃描契變作業菜單按鈕
    Menu2=(By.XPATH,'/html/body/div/div[2]/div/ul/li[1]')
    # 契變撤銷作業菜單按鈕
    Menu_preser_reject=(By.XPATH,'/html/body/div/div[2]/div/ul/li[10]')
    preser_reject_apply=(By.PARTIAL_LINK_TEXT,'契變撤銷申請')
    preser_reject_agree=(By.PARTIAL_LINK_TEXT,'契變撤銷覆核')

    # 保全處理_自動核保
    preser_autocomplete=(By.XPATH,'/html/body/div/div[2]/div/ul/li[28]/a')
    # 保全處理_契變審核作業
    Review=(By.XPATH,'/html/body/div/div[2]/div/ul/li[6]/a')

    # 客戶/保單號
    PublicWorkPoolQueryGrid3r0=(By.ID,'PublicWorkPoolQueryGrid3r0')

    # 我的任務池內保單
    PrivateWorkPoolGridChk0=(By.ID,'PrivateWorkPoolGridChk0')
    # 自動核保按鈕
    Autosubmit=(By.XPATH,'/html/body/form/a[1]')

    menuButton4=(By.PARTIAL_LINK_TEXT,'系統管理')
# -------設定當前系統時間--------------------------------------------------------------------------
    Menu3=(By.XPATH,'/html/body/div/div[2]/div/ul/li[3]/a')
    # Menu3=(By.XPATH,'/html/body/div/div[2]/div/ul/li[4]/a')
    # Menu3=(By.CSS_SELECTOR,'')
    timesetup=(By.ID,'TimeConf')


    SetTime=(By.ID,'SetTime')
    butSubmit=(By.XPATH,'/html/body/center/input')
# -------------------------------------------------------------------------------------------------
    ### 批次處理任務執行
    Menu4=(By.XPATH,'/html/body/div/div[2]/div/ul/li[1]/a')
# -------------------------------------------------------------------------------------------------

    menuButton5=(By.PARTIAL_LINK_TEXT,'理賠案件')

    ### 受理作業
    acceptWork=(By.XPATH,'/html/body/div/div[3]/div/dl[1]/dt')
    ### 紙本受理
    # acceptScan=(By.XPATH,'/html/body/div/div[3]/div/dl[1]/dd[1]/a')
    # acceptScan=(By.XPATH,'/html/body/div/div[3]/div/dl[1]/dd[2]/a')
    acceptScan=(By.CSS_SELECTOR,'a.innerA')
    # /html/body/div/div[3]/div/dl[1]/dd[2]/a

# --------保單列印---------
    PrintSignPage=(By.PARTIAL_LINK_TEXT,'保單列印')

# --------人工核保---------
    innerA=(By.PARTIAL_LINK_TEXT,'人工核保')
    # innerA=(By.CSS_SELECTOR,'.innerA')
    # ArticalConfirm=(By.ID,'huadong1')

# ---------自動核保頁面(需參考"U_AutoConfirmPage"內的方法，此處只有導航到自動核保的menu)---------
    AutoConfirm=(By.PARTIAL_LINK_TEXT,'自動核保')

    NewIInput=(By.PARTIAL_LINK_TEXT,'新單錄入')
# ---------新單覆核頁面-----------------------------------
    Newconfirm=(By.PARTIAL_LINK_TEXT,'新單覆核')

    # 查詢日期
    confirm_date=(By.ID,'PublicWorkPoolQueryGrid3r0')

    # 保單號碼
    confirm_PoliceNumber=(By.ID,'PublicWorkPoolQueryGrid1r0')

    # 查詢按鈕
    confirm_queryBtn=(By.ID,'publicSearch')

    # 池內第一組保單
    confirm_pool=(By.ID,'PublicWorkPoolGridSel0')

    # 申請按鈕
    confirm_apply=(By.ID,'riskbutton')

    # 掃描案號
    ScanCode=(By.ID,'PrivateWorkPoolQueryGrid1r0')
    # 查詢按鈕
    ScanSearch=(By.ID,'privateSearch')

    # 個人工作池內保單
    confirm_per_pool=(By.ID,'PrivateWorkPoolGridSel0')
    confirm_per_pool2=(By.ID,'PrivateWorkPoolGridSel1')
    confirm_per_pool3=(By.ID,'PrivateWorkPoolGridSel2')

    # 覆核按鈕
    confirm_confirmBtn=(By.ID,'Donextbutton4')

# ---------新單查詢頁面-----------------------------------
    NewQuery=(By.PARTIAL_LINK_TEXT,'新單査詢')

    # 受理收件日期
    AcceptDate =(By.ID,'AcceptDate')
    
    # 要保日期
    queryDate =(By.ID,'PolApplyDate')

    # 客戶姓名
    queryName =(By.ID,'CustomernoName')

    # 客戶身分證
    queryID=(By.ID,'CustomernoIDNo')

    # 保單號碼
    ContNo=(By.ID,'ContNo')

    # 契變案號
    EdorAcceptNo=(By.ID,'EdorAcceptNo')

    # 保單詳細資訊前面的圈圈
    queryBtn=(By.PARTIAL_LINK_TEXT,'查詢')

    seleOne =(By.ID,'PolGridSel0')
    poolsele=(By.ID,'PolDetailedGridSel0')

    # 保單詳細資訊
    insuranceNb =(By.ID,'PolDetailedGrid1r0')
    

    # 真實姓名
    RealName=(By.ID,'PolDetailedGrid2r0')

    # 保費 
    # RealPrice=(By.ID,'LjsFeeClassGrid8r0')
    # 保費總額
    # RealPrice=(By.ID,'AfterDiscountSumFee')
    # 短溢繳
    # RealPrice=(By.ID,'DaSumFee')
    # 應收總保費
    RealPrice=(By.ID,'LjsSumFee')
    # RealPrice=(By.CSS_SELECTOR,'input#LjsFeeClassGrid8r0.mulreadonly')
    

    # 繳費資訊查詢
    Button07=(By.ID,'Button07')

# ---------簽發保單頁面-----------------------------------
    SignPage=(By.PARTIAL_LINK_TEXT,'簽發保單')
# ---------簽收保單頁面-----------------------------------
    SignBackPage=(By.PARTIAL_LINK_TEXT,'簽收回條')

# ---------立帳頁面-----------------------------------
    Cashier =(By.PARTIAL_LINK_TEXT,'出納')
    # 現金/匯款/劃撥立帳
    Account2=(By.ID,'1')
    Account3=(By.PARTIAL_LINK_TEXT,'單筆立帳')
    AccountQ=(By.PARTIAL_LINK_TEXT,'立帳綜合查詢')
    # 查詢按鈕
    queryButton =(By.CSS_SELECTOR,'input.cssButton')

# ---------保費頁面(需參考"PremiumPage"內的方法)-----------------------------------
    Premium =(By.PARTIAL_LINK_TEXT,'保費')
    Premium1=(By.PARTIAL_LINK_TEXT,'單筆自繳建檔')
    PremiumCheck=(By.PARTIAL_LINK_TEXT,'單筆自繳審核')
    Inform=(By.PARTIAL_LINK_TEXT,'通知書查詢')
# ---------銀行轉帳頁面-----------------------------------
    AuthBook=(By.PARTIAL_LINK_TEXT,'授權書建檔')
    AuthBook_Delivery=(By.PARTIAL_LINK_TEXT,'核印送件作業')
    AuthBook_Approval=(By.PARTIAL_LINK_TEXT,'核印送核作業')
    AuthBook_Return=(By.PARTIAL_LINK_TEXT,'核印回件作業')


# ---------綜合查詢-----------------------------------
    # 綜合查詢
    menuButton6=(By.PARTIAL_LINK_TEXT,'綜合查詢')
    # 個人保單查詢
    Personal_Query=(By.XPATH,'/html/body/div/div[2]/div/ul/li[5]/a')
    # 保單號碼
    PrtNo=(By.ID,'PrtNo')
    # 查詢按鈕
    Qbtn=(By.XPATH,'/html/body/form/a[1]')
    # 保單資訊(水池)
    PolGridSel0=(By.ID,'PolGridSel0')
    # 保單明細
    detail=(By.XPATH,'/html/body/form/p/a')
    # 交費對應日
    PolGrid17r0=(By.ID,'PolGrid17r0')

    # ====保全查詢============================
    Preservation=(By.XPATH,'/html/body/div/div[2]/div/ul/li[10]/a')
    Preser_QueryBtn=(By.XPATH,'/html/body/form/input[1]')
    ContNo=(By.ID,'ContNo')
    Preservation_Date=(By.ID,'PolGrid5r0')
    # 保全明細按鈕
    PreserListBtn=(By.XPATH,'/html/body/form/input[2]')
    # 補退費訊息
    divSumGetPayInfo=(By.ID,'divSumGetPayInfo')
# ---------綜合查詢頁面_保全查詢----------------------------------------------------------------
    def mouse_To_QueryPage_For_PreservationQuery(self,Policy_number):
        self.switch_default()
        self.switch_frame('head')
        for i in range(5):
            self.find_element(*self.menuButton_next).click()
            print("點擊左邊箭頭"+ str(i) +"次")
        ### 綜合查詢
        self.find_element(*self.menuButton6).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        ### 保全查詢
        self.find_element(*self.Preservation).click()
        self.switch_default()
        self.switch_frame("fraInterface")
        self.find_element(*self.ContNo).send_keys(Policy_number)
        self.find_element(*self.Preser_QueryBtn).click()
        Preser_Date=self.find_element(*self.Preservation_Date).get_attribute('value')
        return Preser_Date

# ---------綜合查詢頁面_保全查詢(契變案號查詢)----------------------------------------------------------------
    def mouse_To_QueryPage_For_PreserEdorNumber(self,EdorNumber):
        self.switch_default()
        self.switch_frame('head')
        for i in range(5):
            self.find_element(*self.menuButton_next).click()
            print("點擊左邊箭頭"+ str(i) +"次")
        ### 綜合查詢
        self.find_element(*self.menuButton6).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        ### 保全查詢
        self.find_element(*self.Preservation).click()
        self.switch_default()
        self.switch_frame("fraInterface")
        self.find_element(*self.EdorAcceptNo).send_keys(EdorNumber)
        self.find_element(*self.Preser_QueryBtn).click()
        # 選擇水池內保單
        self.find_element(*self.seleOne).click()
        # 點擊保全明細按鈕
        self.find_element(*self.PreserListBtn).click()

        # 跳轉窗口
        self.switch_window(1)
        self.switch_default()
        self.switch_frame("fraInterface")
        print("測試8:"+self.driver.title)
        # 移動到補退費訊息 divSumGetPayInfo
        element=self.find_element(*self.divSumGetPayInfo)
        self.driver.execute_script("arguments[0].scrollIntoView(true);", element)
        self.take_screenshot()


        # Preser_Date=self.find_element(*self.Preservation_Date).get_attribute('value')
        # return Preser_Date

# ---------綜合查詢頁面_查詢續期的交費對應日----------------------------------------------------------------
    def mouse_To_QueryPage_For_RenewPremiumDate(self,Policy_number):
        self.switch_default()
        self.switch_frame('head')
        for i in range(5):
            self.find_element(*self.menuButton_next).click()
            print("點擊左邊箭頭"+ str(i) +"次")
        ### 綜合查詢
        self.find_element(*self.menuButton6).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        ### 個人保單査詢
        self.find_element(*self.Personal_Query).click()
        for i in range(len(Policy_number)):
            self.switch_default()
            self.switch_frame("fraInterface")
            ### 輸入保單號碼
            self.find_element(*self.PrtNo).clear()
            self.find_element(*self.PrtNo).send_keys(Policy_number[i])
            self.find_element(*self.Qbtn).click()
            self.find_element(*self.PolGridSel0).click()
            ### 保單明細按鈕 
            self.find_element(*self.detail).click()
            ### 跳轉視窗
            self.switch_window(1)
            ### 確定跳轉視窗的名稱
            print(self.driver.title)
            self.switch_default()
            self.switch_frame("fraInterface")
            ### 得到續期的交費對應日
            renewable_premium_Date=self.find_element(*self.PolGrid17r0).get_attribute('value')
            print("交費對應日:"+renewable_premium_Date)
            # 跳回原本窗口
            self.switch_window_back()
            ### 確定跳轉視窗的名稱
            print(self.driver.title)
        # self.switch_default()
        # self.switch_frame("fraInterface")
        return renewable_premium_Date

    def mouse_To_QueryPage_For_RenewPremiumDate_for_Preservation(self,Policy_number):
        self.switch_default()
        self.switch_frame('head')
        for i in range(5):
            self.find_element(*self.menuButton_next).click()
            print("點擊左邊箭頭"+ str(i) +"次")
        ### 綜合查詢
        self.find_element(*self.menuButton6).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        ### 個人保單査詢
        self.find_element(*self.Personal_Query).click()
        self.switch_default()
        self.switch_frame("fraInterface")
        ### 輸入保單號碼
        self.find_element(*self.PrtNo).clear()
        self.find_element(*self.PrtNo).send_keys(Policy_number)
        self.find_element(*self.Qbtn).click()
        self.find_element(*self.PolGridSel0).click()
        ### 保單明細按鈕 
        self.find_element(*self.detail).click()
        ### 跳轉視窗
        self.switch_window(1)
        ### 確定跳轉視窗的名稱
        print(self.driver.title)
        self.switch_default()
        self.switch_frame("fraInterface")
        ### 得到續期的交費對應日
        renewable_premium_Date=self.find_element(*self.PolGrid17r0).get_attribute('value')
        print("交費對應日:"+renewable_premium_Date)
        # 移到交費對應日的前一個月
        test_data = datetime.datetime.strptime(renewable_premium_Date, '%Y-%m-%d')
        date1=test_data.replace(day=1)
        last_mon=date1-datetime.timedelta(days=21)
        last_mon2=last_mon.replace(day=10)
        last_monDate=last_mon2.strftime("%Y-%m-%d")
        print("交費對應日的前一個月的10號:"+last_monDate)
        # 跳回原本窗口
        self.switch_window_back()
        ### 確定跳轉視窗的名稱
        print(self.driver.title)
        # self.switch_default()
        # self.switch_frame("fraInterface")
        return last_monDate


    def mouse_To_SystemTime(self,personinfo):
        self.switch_default()
        self.switch_frame('head')
        for i in range(2):
            self.find_element(*self.menuButton_next).click()
            print("點擊左邊箭頭"+ str(i) +"次")
        self.find_element(*self.menuButton4).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        # 設定當前系統時間
        self.find_element(*self.Menu3).click()
        self.switch_default()
        self.switch_frame("fraInterface")
        # 設定系統時間
        # 使用javascript控制日歷
        js='document.getElementById("TimeConf").removeAttribute("readonly");'
        self.driver.execute_script(js)
        # js_value='document.getElementById("TimeConf").value = {}' .format("2021-10-17")
        # print(js_value)
        # self.driver.execute_script(js_value)
        self.find_element(*self.timesetup).send_keys(personinfo["PolAppntDate"])
        self.find_element(*self.SetTime).click()
        # 切換到彈窗
        alert = self.driver.switch_to.alert
        # 列印彈窗顯示的資訊:Alert Message
        print(alert.text) 
        # 接受
        alert.accept()
        time.sleep(1)

    def mouse_To_SystemTime_for_preservation(self,Preser_Date):
        self.switch_default()
        self.switch_frame('head')
        self.find_element(*self.menuButton4).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        # 設定當前系統時間
        self.find_element(*self.Menu3).click()
        self.switch_default()
        self.switch_frame("fraInterface")
        # 設定系統時間
        # 使用javascript控制日歷
        js='document.getElementById("TimeConf").removeAttribute("readonly");'
        self.driver.execute_script(js)
        # js_value='document.getElementById("TimeConf").value = {}' .format("2021-10-17")
        # print(js_value)
        # self.driver.execute_script(js_value)
        self.find_element(*self.timesetup).send_keys(Preser_Date)
        self.find_element(*self.SetTime).click()
        # 切換到彈窗
        alert = self.driver.switch_to.alert
        # 列印彈窗顯示的資訊:Alert Message
        print(alert.text) 
        # 接受
        alert.accept()
        time.sleep(1)

    
# ---------系統管理頁面(移動到批次處理任務執行)-----------------------------------
    def mouse_To_SystemTaskSetup(self):
        self.switch_default()
        self.switch_frame('head')
        self.find_element(*self.menuButton4).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        # 批次處理任務執行
        self.find_element(*self.Menu4).click()
        self.switch_default()
        self.switch_frame("fraInterface")

        
# ---------理賠頁面(移動到理賠)-----------------------------------
    def mouse_To_claim(self):
        self.switch_default()
        self.switch_frame('head')
        self.find_element(*self.menuButton5).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        claim = self.find_elements(*self.Menu)
        ### 個險
        claim1 = claim[0]
        ActionChains(self.driver).move_to_element(claim1).click().perform()
        
        ### 受理作業
        self.find_element(*self.acceptWork).click()

        ### 紙本受理 
        acceptScan1 = self.find_elements(*self.acceptScan)
        acceptScan2 = acceptScan1 [8]
        ActionChains(self.driver).move_to_element(acceptScan2).click().perform()
        self.switch_default()
        self.switch_frame("fraInterface")

# ---------綜合管理頁面(移動到保全案件調撥)-----------------------------------
    def mouse_To_Preservation_return(self,Policy_number):
        self.switch_default()
        self.switch_frame('head')
        for i in range(2):
            self.find_element(*self.menuButton_next).click()
        print("點擊左邊箭頭"+ str(i) +"次")
        # 點擊綜合管理
        self.find_element(*self.menuButton7).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        # 點擊保全案件調撥
        preser_return=self.find_elements(*self.preser_return)
        preser_return2 = preser_return[2]
        ActionChains(self.driver).move_to_element(preser_return2).click().perform()

        self.switch_default()
        self.switch_frame("fraInterface")

        # 輸入契變崗位
        self.find_element(*self.ChangePost).click()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('0000000019')

        # 輸入保單號碼
        self.find_element(*self.ContNo).send_keys(Policy_number)
        # 查詢按鈕
        self.find_element(*self.preser_return_Btn).click()

        # 選擇個人池
        self.find_element(*self.checkAllPersonalWorkPoolGrid).click()
        
        # 處理人員
        self.find_element(*self.ProcessOperater).send_keys('TEST33')
        self.find_element(*self.codeselect).click()

        # 保存按鈕
        saveBtn=self.find_elements(*self.preser_return_Btn)
        saveBtn2=saveBtn[2]
        ActionChains(self.driver).move_to_element(saveBtn2).click().perform()

        self.switch_default()
        self.switch_frame('head')
        for i in range(2):
            self.find_element(*self.menuButton_prev).click()
        print("點擊左邊箭頭"+ str(i) +"次")

        


# ---------保全處理頁面(移動到無掃描契變作業)-----------------------------------
    def mouse_To_Preservation(self):
        self.switch_default()
        self.switch_frame('head')
        self.find_element(*self.menuButton3).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        # 無掃描契變作業
        self.find_element(*self.Menu2).click()
        self.switch_default()
        self.switch_frame("fraInterface")

# ---------保全處理頁面(移動到契變撤銷作業(申請))-----------------------------------
    def mouse_To_Preservation_rejectApply(self):
        self.switch_default()
        self.switch_frame('head')
        self.find_element(*self.menuButton3).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        # 契變撤銷作業
        self.find_element(*self.Menu_preser_reject).click()
        # 契變撤銷申請
        self.find_element(*self.preser_reject_apply).click()
        self.switch_default()
        self.switch_frame("fraInterface")

# ---------保全處理頁面(移動到契變撤銷作業(覆核))-----------------------------------
    def mouse_To_Preservation_rejectAgree(self):
        self.switch_default()
        self.switch_frame('head')
        self.find_element(*self.menuButton3).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        # 契變撤銷作業
        self.find_element(*self.Menu_preser_reject).click()
        # 契變撤銷覆核
        self.find_element(*self.preser_reject_agree).click()
        self.switch_default()
        self.switch_frame("fraInterface")

# ---------保全處理頁面(移動到自動核保)-----------------------------------
    def mouse_To_Preservation_Autocomplete(self,Policy_number):
        self.switch_default()
        self.switch_frame('head')
        self.find_element(*self.menuButton3).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        # 自動核保
        self.find_element(*self.preser_autocomplete).click()
        self.switch_default()
        self.switch_frame("fraInterface")
        # 輸入保單號碼
        self.find_element(*self.PublicWorkPoolQueryGrid3r0).send_keys(Policy_number)
        # 點查詢按鈕
        self.find_element(*self.confirm_queryBtn).click()
        # 選擇池內保單
        self.find_element(*self.confirm_pool).click()
        # 點申請按鈕
        self.find_element(*self.confirm_apply).click()
        # 選擇我的任務池內保單
        self.find_element(*self.PrivateWorkPoolGridChk0).click()
        # 點擊自動核保
        self.find_element(*self.Autosubmit).click()

# ---------保全處理頁面(移動到契變審核作業)-----------------------------------
    def mouse_To_Preservation_Review(self):
        self.switch_default()
        self.switch_frame('head')
        self.find_element(*self.menuButton3).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        # 契變審核作業
        self.find_element(*self.Review).click()
        self.switch_default()
        self.switch_frame("fraInterface")

# -----------------移動到"新單錄入"-------------------------------
    def mouse_To_NewIInput(self,Scan):
        self.switch_default()
        self.switch_frame('head')
        self.find_element(*self.menuButton).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        goon = self.find_elements(*self.Menu)
        # 個人保單
        goon1 = goon[0]
        ActionChains(self.driver).move_to_element(goon1).click().perform()
        # 新單錄入
        self.find_element(*self.NewIInput).click()
        self.switch_default()
        self.switch_frame("fraInterface")

        # 掃描案號
        self.find_element(*self.ScanCode).send_keys(Scan)
        # 查詢按鈕
        self.find_element(*self.ScanSearch).click()
        # 第一單
        self.find_element(*self.confirm_per_pool).click()
        # # 第二單
        # self.find_element(*self.confirm_per_pool2).click()
        # # 第三單
        # self.find_element(*self.confirm_per_pool3).click()




# -----------------移動到"新單覆核"-------------------------------
    def mouse_To_NewConform(self):
        self.switch_default()
        self.switch_frame('head')
        self.find_element(*self.menuButton).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        goon = self.find_elements(*self.Menu)
        # 個人保單
        goon1 = goon[0]
        ActionChains(self.driver).move_to_element(goon1).click().perform()
        # 新單覆核
        self.find_element(*self.Newconfirm).click()
        self.switch_default()
        self.switch_frame("fraInterface")

# -----------------移動到"新單查詢"------------------------------- 
    def mouse_To_NewQuery(self):
        self.switch_default()
        self.switch_frame('head')
        self.find_element(*self.menuButton).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        goon = self.find_elements(*self.Menu)
        # 個人保單
        goon1 = goon[0]
        ActionChains(self.driver).move_to_element(goon1).click().perform()
        # 新單查詢
        self.find_element(*self.NewQuery).click()
        self.switch_default()
        self.switch_frame("fraInterface")
    # 移動到新單查詢頁面內的"繳費資訊查詢"按鈕   
    def mouse_To_button07(self):
        # 新單查詢頁面內的"繳費資訊查詢"按鈕
        self.find_element(*self.Button07).click()

        # 換視窗，裡面的1是變數可以決定更換為第幾個視窗
        self.switch_window(1)
        self.switch_default()
        self.switch_frame("fraInterface")


# -----------------移動到"自動核保"--------------------------------
    def mouse_To_Autoconfirm(self):
        self.switch_default()
        self.switch_frame('head')
        self.find_element(*self.menuButton).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        goon = self.find_elements(*self.Menu)
        # 個人保單
        goon1 = goon[0]
        ActionChains(self.driver).move_to_element(goon1).click().perform()
        self.find_element(*self.AutoConfirm).click()
        self.switch_default()
        self.switch_frame("fraInterface")

# -----------------移動到人工核保-------------------------------- 
    def mouse_To_Articalconfirm(self):
        self.switch_default()
        self.switch_frame('head')
        self.find_element(*self.menuButton).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        goon = self.find_elements(*self.Menu)
        # 點擊個人保單
        goon1 = goon[0]
        ActionChains(self.driver).move_to_element(goon1).click().perform()

        # 選擇人工核保
        menuinnerA = self.find_elements(*self.innerA)
        menuinnerA2 = menuinnerA[1]
        ActionChains(self.driver).move_to_element(menuinnerA2).click().perform()
        self.switch_default()
        self.switch_frame("fraInterface")

# -----------------移動到"單筆立帳"-------------------------------- 
    def mouse_To_Account(self):
        self.switch_default()
        self.switch_frame('head')
        self.find_element(*self.Cashier).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        goon = self.find_elements(*self.Menu)
        # 立帳作業
        goon1 = goon[1]
        ActionChains(self.driver).move_to_element(goon1).click().perform()
        self.find_element(*self.Account2).click()
        self.find_element(*self.Account3).click()
        self.switch_default()
        self.switch_frame("fraInterface")

# -----------------移動到"立帳綜合查詢"--------------------------------
    def mouse_To_AccountQuery(self):
        self.switch_default()
        self.switch_frame('head')
        self.find_element(*self.Cashier).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        goon = self.find_elements(*self.Menu)
        # 立帳作業
        goon1 = goon[1]
        ActionChains(self.driver).move_to_element(goon1).click().perform()
        self.find_element(*self.AccountQ).click()
        self.switch_default()
        self.switch_frame("fraInterface")

# -----------------移動到銀行轉帳--------------------------------
    def mouse_To_BankTransfer_AuthBook(self):
        self.switch_default()
        self.switch_frame('head')
        # 移動到保費
        self.find_element(*self.Premium).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        goon = self.find_elements(*self.Menu)
        goon12=goon[0]
        # 移動到銀行轉帳
        ActionChains(self.driver).move_to_element(goon12).click().perform()
        # 移動到授權書建檔
        self.find_element(*self.AuthBook).click()
        self.switch_default()
        self.switch_frame("fraInterface")

    def mouse_To_BankTransfer_Delivery(self):
        self.switch_default()
        self.switch_frame('head')
        # 移動到保費
        self.find_element(*self.Premium).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        goon = self.find_elements(*self.Menu)
        goon12=goon[0]
        # 移動到銀行轉帳
        ActionChains(self.driver).move_to_element(goon12).click().perform()
        # 移動到授權書建檔
        self.find_element(*self.AuthBook_Delivery).click()
        self.switch_default()
        self.switch_frame("fraInterface")

    def mouse_To_BankTransfer_Approval(self):
        self.switch_default()
        self.switch_frame('head')
        # 移動到保費
        self.find_element(*self.Premium).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        goon = self.find_elements(*self.Menu)
        goon12=goon[0]
        # 移動到銀行轉帳
        ActionChains(self.driver).move_to_element(goon12).click().perform()
        # 移動到授權書建檔
        self.find_element(*self.AuthBook_Approval).click()
        self.switch_default()
        self.switch_frame("fraInterface")
        
    def mouse_To_BankTransfer_Return(self):
        self.switch_default()
        self.switch_frame('head')
        # 移動到保費
        self.find_element(*self.Premium).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        goon = self.find_elements(*self.Menu)
        goon12=goon[0]
        # 移動到銀行轉帳
        ActionChains(self.driver).move_to_element(goon12).click().perform()
        # 移動到授權書建檔
        self.find_element(*self.AuthBook_Return).click()
        self.switch_default()
        self.switch_frame("fraInterface")

# -----------------移動到續期作業-------------------------------- 
    def mouse_To_Renew(self):
        self.switch_default()
        self.switch_frame('head')
        # 移動到保費
        self.find_element(*self.Premium).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        goon = self.find_elements(*self.Menu)
        # 續期作業
        goon1 = goon[7]
        ActionChains(self.driver).move_to_element(goon1).click().perform()
        self.find_element(*self.Inform).click()
        self.switch_default()
        self.switch_frame("fraInterface")

# -----------------移動到單筆自繳-------------------------------- 
    def mouse_To_Premium(self):
        self.switch_default()
        self.switch_frame('head')
        # 移動到保費
        self.find_element(*self.Premium).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        goon = self.find_elements(*self.Menu)
        # 單筆自繳
        goon1 = goon[3]
        ActionChains(self.driver).move_to_element(goon1).click().perform()
        self.find_element(*self.Premium1).click()
        self.switch_default()
        self.switch_frame("fraInterface")

# -----------------移動到單筆自繳審核--------------------------------
    def mouse_To_PremiumCheck(self):
        self.switch_default()
        self.switch_frame('head')
        # 移動到保費
        self.find_element(*self.Premium).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        goon = self.find_elements(*self.Menu)
        # 單筆自繳
        goon1 = goon[3]
        ActionChains(self.driver).move_to_element(goon1).click().perform()
        self.find_element(*self.PremiumCheck).click()
        self.switch_default()
        self.switch_frame("fraInterface")
    
# -----------------# 移動到簽發保單---------------------------------
    def mouse_To_SignPage(self):
        self.switch_default()
        self.switch_frame('head')
        self.find_element(*self.menuButton).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        goon = self.find_elements(*self.Menu)
        # 個人保單
        goon1 = goon[0]
        ActionChains(self.driver).move_to_element(goon1).click().perform()
        self.find_element(*self.SignPage).click()
        self.switch_default()
        self.switch_frame("fraInterface")

# -----------------移動到綜合列印的承保列印---------------------------
    def mouse_To_PrintSignPage(self):
        self.switch_default()
        self.switch_frame('head')
        for i in range(5):
            self.find_element(*self.menuButton_next).click()
            print("點擊右邊箭頭"+ str(i) +"次")
        # 綜合列印
        self.find_element(*self.menuButton2).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        goon = self.find_elements(*self.Menu)
        # 承保列印
        goon1 = goon[1]
        ActionChains(self.driver).move_to_element(goon1).click().perform()
        self.find_element(*self.PrintSignPage).click()
        self.switch_default()
        self.switch_frame("fraInterface")

# -----------------移動到簽收保單---------------------------
    def mouse_To_SignBackPage(self):
        self.switch_default()
        self.switch_frame('head')
        for i in range(5):
            self.find_element(*self.menuButton_prev).click()
            print("點擊左邊箭頭"+ str(i) +"次")
        self.find_element(*self.menuButton).click()
        self.switch_default()
        self.switch_frame("fraMenu")
        goon = self.find_elements(*self.Menu)
        # 個人保單
        goon1 = goon[0]
        ActionChains(self.driver).move_to_element(goon1).click().perform()
        self.find_element(*self.SignBackPage).click()
        self.switch_default()
        self.switch_frame("fraInterface")

# -----------------新單覆核---------------------------
    def New_Conform(self,personinfo,Policy_number):
        # 點擊後輸入日期
        self.find_element(*self.confirm_date).click()
        self.find_element(*self.confirm_date).send_keys(personinfo["PolAppntDate"])

        time.sleep(1)
        # 輸入保單號碼 PublicWorkPoolQueryGrid1r0
        self.find_element(*self.confirm_PoliceNumber).send_keys(Policy_number)

        # 點擊查詢
        self.find_element(*self.confirm_queryBtn).click()
        time.sleep(1)

        # 選擇公共池內的保單
        self.find_element(*self.confirm_pool).click()
        
        # 申請
        self.find_element(*self.confirm_apply).click()
        time.sleep(1)

        # 選擇個人工作池內的保單
        self.find_element(*self.confirm_per_pool).click()

        # 移動到第二視窗
        t0=self.driver.title
        print("目前窗口1"+t0)
        
        # 轉換到新單覆核窗口
        self.switch_window(2)
        t1=self.driver.title
        print("目前窗口2:"+t1)

        # 若窗口為"資訊回饋"則再跳轉一次
        if (t1=="資訊回饋"):
            # # 轉換到新單覆核窗口
            self.switch_window(3)
            t2=self.driver.title
            print("目前窗口3:"+t2)
        
        # 此處是因為網頁有跳出"請求失敗"，為解決此訊息而寫，未來可以mark掉
        # self.switch_to_alert_accept()
        
        # # 新單覆核完畢 Donextbutton4
        self.switch_default()
        self.switch_frame("fraInterface")
        # 點擊覆核按鈕
        self.find_element(*self.confirm_confirmBtn).click()
        # print("點擊成功")
        time.sleep(2)


# -----------------新單查詢方法-----------------------

    # 受理收件日期
    def input_QueryPage_acceptDate(self,personinfo):
        self.find_element(*self.AcceptDate).send_keys(personinfo["AcceptDate"])

    # 要保日期
    def input_QueryPage_Date(self,personinfo):
        self.find_element(*self.queryDate).send_keys(personinfo["PolAppntDate"])
    # 客戶姓名
    def input_QueryPage_Name(self,personinfo):
        self.find_element(*self.queryName).send_keys(personinfo["AppntName"])
    # 客戶身分證
    def input_QueryPage_ID(self,personinfo):
        self.find_element(*self.queryID).send_keys(personinfo["AppntIDNo"])

    # 保單號碼
    def input_QueryPage_PolicyNB(self,personinfo):
        self.find_element(*self.ContNo).send_keys(personinfo["Policy_number"])

    # 查詢按鈕
    def click_QueryPage_Button(self):
        self.find_element(*self.queryButton).click()
    # 選擇保單詳細資訊第一欄的圈圈
    def check_QueryPage_seleOne(self):
        self.find_element(*self.seleOne).click()
        self.find_element(*self.poolsele).click()
    # 得到保單號碼
    def getNumbertext(self):
        number=self.find_element(*self.insuranceNb).get_attribute('value')
        return number
    #得到姓名 
    def getName(self):
        RealName1 =self.find_element(*self.RealName).get_attribute('value')
        return RealName1
    # 得到保費
    def getRealPrice(self):
        RealPrice = self.find_element(*self.RealPrice).get_attribute('value')
        return RealPrice



        
