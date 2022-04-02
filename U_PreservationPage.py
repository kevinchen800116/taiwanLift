
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from selenium.common.exceptions import TimeoutException
import time

# 現在的位置：保全處理-->契變撤銷作業-->契變撤銷申請
class Preservation_rejectPage():
    # --------------申請---------------
    # 輸入保單號碼
    PrivateWorkPoolQueryGrid3r0=(By.ID,'PrivateWorkPoolQueryGrid3r0')
    # 申請頁面的查詢按鈕
    privateSearch=(By.ID,'privateSearch')
    # 水池內第一個選項
    PrivateWorkPoolGridSel0=(By.ID,'PrivateWorkPoolGridSel0')
    # 保全撤銷
    rejectBtn=(By.CSS_SELECTOR,'a.button')
    # 撤銷原因
    MCanclReason=(By.ID,'MCanclReason')
    # 選擇框
    codeselect=(By.ID,'codeselect')
    # 保存
    insertButton=(By.ID,'insertButton')
    # ---------------覆核----------------
    PublicWorkPoolQueryGrid3r0=(By.ID,'PublicWorkPoolQueryGrid3r0')
    # 覆核頁面的查詢按鈕
    publicSearch=(By.ID,'publicSearch')
    # 選擇水池內第一個案件 PrivateWorkPoolGridSel0
    PublicWorkPoolGridSel0=(By.ID,'PublicWorkPoolGridSel0')
    # 私人池內第一個案件
    PrivateWorkPoolGridSel0=(By.ID,'PrivateWorkPoolGridSel0')
    # 契變撤銷覆核
    submit=(By.PARTIAL_LINK_TEXT,'契變撤銷覆核')
    # 取件按鈕
    riskbutton=(By.ID,'riskbutton')
    # 覆核結論
    ApproveReason=(By.ID,'ApproveReason')
    # 覆核意見
    ApproveReasonContent=(By.ID,'ApproveReasonContent')
    # 覆核確認
    confirmButton=(By.ID,'confirmButton')
    # 人工核保按鈕
    ManuUW=(By.ID,'ManuUW')
    # 契變案號
    EdorMainGrid3r0=(By.ID,'EdorMainGrid3r0')
    # 客戶承保結論變更 
    indButton7=(By.ID,'indButton7')
    # 保存按鈕
    button01=(By.ID,'button01')
    # 除外資訊div
    divLCPolSpec=(By.ID,'divLCPolSpec')
    # 加費資訊div
    divLCPolApp=(By.ID,'divLCPolApp')
    # 加費資訊
    PolAddGridaddOne=(By.ID,'PolAddGridaddOne')
    # 險種名稱
    PolAddGrid2r0=(By.ID,'PolAddGrid2r0')
    # 加費方式
    PolAddGrid5r0=(By.ID,'PolAddGrid5r0')
    # 評點比例
    PolAddGrid8r0=(By.ID,'PolAddGrid8r0')
    # 加費起日
    PolAddGrid13r0=(By.ID,'PolAddGrid13r0')
    # 折扣前加費金額
    PolAddGrid9r0=(By.ID,'PolAddGrid9r0')
    # 選擇次標準體
    OldRislPlanGrid10r0=(By.ID,'OldRislPlanGrid10r0')

    # 疾病疾病的加號(+)
    AddDiseaseGridaddOne=(By.ID,'AddDiseaseGridaddOne')

    # 打勾險種名稱
    PolAddGridSel0=(By.ID,'PolAddGridSel0')
    # 打勾險種名稱
    AddDiseaseGridChk0=(By.ID,'AddDiseaseGridChk0')

    # 選擇商品結論變更的保單
    OldRislPlanGridSel0=(By.ID,'OldRislPlanGridSel0')

    # 確定按鈕
    btnConfirm=(By.ID,'btnConfirm')
    # 返回按鈕
    Back=(By.XPATH,'/html/body/form/div[4]/div[5]/div[2]/table/tbody/tr/td/div[1]/input[2]')

    # 人工核保的核保意見
    UWIdea=(By.ID,'UWIdea')
    # 人工核保確認按鈕
    Arti_SubmitBtn=(By.ID,'SubmitBtn')


    # 申請
    def input_PreservationPage_reject_apply(self,Policy_number):
        self.find_element(*self.PrivateWorkPoolQueryGrid3r0).send_keys(Policy_number)
        self.find_element(*self.privateSearch).click()
        self.find_element(*self.PrivateWorkPoolGridSel0).click()
        self.find_element(*self.rejectBtn).click()
        # 跳轉視窗
        self.switch_window(1)
        # 確定跳轉視窗的名稱
        print(self.driver.title)
        # 轉換網頁內容
        self.switch_default()
        self.switch_frame("fraInterface")
        # 點擊撤銷原因
        self.find_element(*self.MCanclReason).click()
        b=self.find_element(*self.codeselect)
        ### 選擇撤銷原因 A保戶撤銷 B人工錯誤
        Select(b).select_by_value('B')
        # self.find_element(*self.RgtSourcesType).send_keys(Keys.ENTER)
        self.find_element(*self.insertButton).click()
        self.switch_window_back()
        

    # 覆核
    def input_PreservationPage_reject_agree(self,Policy_number):
        self.find_element(*self.PublicWorkPoolQueryGrid3r0).send_keys(Policy_number)
        # 點擊查詢
        self.find_element(*self.publicSearch).click()

        # # # 選擇公共水池案件
        # self.find_element(*self.PublicWorkPoolGridSel0).click()
        # # # 點擊取件
        # self.find_element(*self.riskbutton).click()

        # ##選擇私人水池案件
        self.find_element(*self.PrivateWorkPoolGridSel0).click()
        # ##點擊契變撤銷覆核 
        self.find_element(*self.submit).click()


        # 跳轉到目標視窗
        self.switch_targetWindow("契變撤銷覆核處理作業")

        # # 獲得當前所有開啟的視窗的控制程式碼
        # sreach_windows = self.driver.current_window_handle
        # all_handles = self.driver.window_handles

        # for handle in all_handles:
        #     if (handle != sreach_windows):
        #         self.driver.switch_to.window(handle)
        #         print("其他窗口:"+self.driver.title)
        #         if(self.driver.title=="契變撤銷覆核處理作業"):
        #             break
        #     else:
        #         print('當前頁面title：%s'%self.driver.title)


        # 轉換網頁內容
        self.switch_default()
        self.switch_frame("fraInterface")
        # 點擊核保結論
        self.find_element(*self.ApproveReason).click()
        b=self.find_element(*self.codeselect)
        ### 選擇結論 1同意
        Select(b).select_by_value('1')
        # 輸入覆核意見
        self.find_element(*self.ApproveReasonContent).send_keys("ok")
        self.find_element(*self.confirmButton).click()
        self.switch_window_back()

# 現在的位置：保全處理-->契變審核作業
    def Preservation_confirm(self,Policy_number,last_monDate):
        # 輸入保單號碼
        self.find_element(*self.PublicWorkPoolQueryGrid3r0).send_keys(Policy_number)
        # 點擊查詢
        self.find_element(*self.publicSearch).click()
        # 選擇公共池內第一個保單
        self.find_element(*self.PublicWorkPoolGridSel0).click()
        # 取件
        self.find_element(*self.riskbutton).click()


        self.switch_window(3)
        self.switch_default()
        self.switch_frame("fraInterface")
        print("測試3:"+self.driver.title)

        # 拍照
        self.take_screenshot()
        self.driver.execute_script("window.scrollTo(0,document.body.scrollHeight);")
        time.sleep(2)
        self.take_screenshot()

        # 人工核保
        self.find_element(*self.ManuUW).click()
        self.switch_window_back()
        self.switch_window(3)
        self.switch_default()
        self.switch_frame("fraInterface")
        print("測試3:"+self.driver.title)
        

        # 客戶承保結論變更 indButton7
        self.find_element(*self.indButton7).click()
        # 跳轉視窗 1保單查詢 2資訊回饋 3人工核保處理作業 4客戶承保結論變更
        self.switch_window(4)
        self.switch_default()
        self.switch_frame("fraInterface")
        print("測試4:"+self.driver.title)
        self.take_screenshot()
        time.sleep(1)

        # 滾動到除外資訊後按鈕後拍照 button01
        element=self.find_element(*self.divLCPolSpec)
        self.driver.execute_script("arguments[0].scrollIntoView(true);", element)
        

        # 滾動到加費資訊
        element=self.find_element(*self.divLCPolApp)
        self.driver.execute_script("arguments[0].scrollIntoView(true);", element)

        # PolAddGridaddOne 加費資訊的加號
        self.find_element(*self.PolAddGridaddOne).click()
        # # PolAddGrid2r0 險種名稱 NUIW2801 輸入後enter
        self.find_element(*self.PolAddGrid2r0).send_keys("N")

        # # 選擇險種 
        # self.find_element(*self.codeselect).click()

        # # PolAddGrid5r0 加費方式
        self.driver.execute_script("PolAddGrid5r0 = document.getElementById('PolAddGrid5r0');" + "PolAddGrid5r0.ondblclick();")
        # self.find_element(*self.PolAddGrid5r0).send_keys("01")

        # # 選擇加費方式
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('01')
        # Select(b).select_by_value('03')
        # self.find_element(*self.codeselect).click()
        # self.find_element(*self.PolAddGrid5r0).send_keys(Keys.ENTER)

        # # PolAddGrid8r0 評點比例 100
        self.find_element(*self.PolAddGrid8r0).send_keys("100")
        

        # # PolAddGrid9r0 折扣前加費金額
        self.driver.execute_script("PolAddGrid9r0 = document.getElementById('PolAddGrid9r0');" + "PolAddGrid9r0.ondblclick();")

        # PolAddGrid11r0 加費開始時間類型
        self.driver.execute_script("PolAddGrid11r0 = document.getElementById('PolAddGrid11r0');" + "PolAddGrid11r0.ondblclick();")
        b=self.find_element(*self.codeselect)
        # 女選1 當期加費  男選0 追溯
        Select(b).select_by_value('0')


        # 加費起日PolAddGrid13r0
        self.find_element(*self.PolAddGrid13r0).clear()
        self.find_element(*self.PolAddGrid13r0).send_keys(last_monDate)


        # 疾病的加號
        self.find_element(*self.AddDiseaseGridaddOne).click()
        self.driver.execute_script("AddDiseaseGrid2r0 = document.getElementById('AddDiseaseGrid2r0');" + "AddDiseaseGrid2r0.ondblclick();")
        b=self.find_element(*self.codeselect)
        # 隨便選
        Select(b).select_by_value('01')
        self.driver.execute_script("AddDiseaseGrid4r0 = document.getElementById('AddDiseaseGrid4r0');" + "AddDiseaseGrid4r0.ondblclick();")
        b=self.find_element(*self.codeselect)
        # 隨便選
        Select(b).select_by_value('010101')

        self.take_screenshot()

        # 打勾險種名稱
        self.find_element(*self.PolAddGridSel0).click()

        # 打勾疾病名稱
        self.find_element(*self.AddDiseaseGridChk0).click()

        # 保存按鈕 button01
        self.find_element(*self.button01).click()

        # 選擇商品結論變更內的保單
        self.find_element(*self.OldRislPlanGridSel0).click()

        # 選擇次標準體 OldRislPlanGrid10r0
        self.driver.execute_script("OldRislPlanGrid10r0 = document.getElementById('OldRislPlanGrid10r0');" + "OldRislPlanGrid10r0.ondblclick();")
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('4')

        self.take_screenshot()


        # 確定按鈕 btnConfirm
        self.find_element(*self.btnConfirm).click()
        # 返回按鈕 
        self.find_element(*self.Back).click()

        # # 跳轉視窗 1保單查詢 2資訊回饋 3人工核保處理作業 4客戶承保結論變更
        self.switch_window_back()
        # 跳到人工核保處理作業
        self.switch_window(3)
        self.switch_default()
        self.switch_frame("fraInterface")
        print("測試5:"+self.driver.title)
        # 取得契變案號.get_attribute('value')
        EdorNumber=self.find_element(*self.EdorMainGrid3r0).get_attribute('value')
        print("契變案號:"+EdorNumber)
        # 點擊客戶承保結論變更
        self.find_element(*self.indButton7).click()
        # # 跳轉視窗 1保單查詢 2資訊回饋 3人工核保處理作業 4客戶承保結論變更
        self.switch_window_back()
        # 跳到客戶承保結論變更
        self.switch_window(5)
        print("測試6:"+self.driver.title)
        self.switch_default()
        self.switch_frame("fraInterface")

        self.take_screenshot()
        # 滾動到加費資訊
        element=self.find_element(*self.divLCPolApp)
        self.driver.execute_script("arguments[0].scrollIntoView(true);", element)
        self.take_screenshot()
        # 滾動到最下方
        self.driver.execute_script("window.scrollTo(0,document.body.scrollHeight);")
        self.take_screenshot()
        # 返回按鈕 
        self.find_element(*self.Back).click()

        # # 跳轉視窗 1保單查詢 2資訊回饋 3人工核保處理作業 4客戶承保結論變更
        self.switch_window_back()
        # 跳到人工核保處理作業
        self.switch_window(3)
        self.switch_default()
        self.switch_frame("fraInterface")
        print("測試7:"+self.driver.title)

        # 處理完畢
        self.driver.execute_script("bqUpReport = document.getElementById('bqUpReport');" + "bqUpReport.ondblclick();")
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('1')

        # 選擇次標準體
        self.driver.execute_script("EdorUWState = document.getElementById('EdorUWState');" + "EdorUWState.ondblclick();")
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('4')

        # 核保意見
        self.find_element(*self.UWIdea).send_keys("ok")
        self.take_screenshot()
        # 確認按鈕
        self.find_element(*self.Arti_SubmitBtn).click()
        return EdorNumber
        

# 現在的位置：保全處理-->無掃描契變作業
class PreservationPage():
    # 保全申請(點了會跳轉頁面)
    ApplyBtn=(By.CSS_SELECTOR,'td.normal')
    # (By.PARTIAL_LINK_TEXT,'保全申請')

    # 搜尋保單號碼
    PrivateWorkPoolQueryGrid3r0=(By.ID,'PrivateWorkPoolQueryGrid3r0')

    # 查詢
    privateSearch=(By.ID,'privateSearch')

    # 水池內第一個選項
    PrivateWorkPoolGridSel0=(By.ID,'PrivateWorkPoolGridSel0')
    PoolAccept=(By.PARTIAL_LINK_TEXT,'保全受理')

    # 輸入保單號碼
    input_PolicyNumber=(By.ID,'OtherNo')

    ### 申請方式(共用選項)
    AppType=(By.ID,'AppType')

    # 選項(申請方式與契變類別共用同一選擇框)
    codeselect=(By.ID,'codeselect')
    # value="01" 同保單服務員(申請方式)
    # value="1" 一般契變(契變類別)
    # value="6" 變更後繳別 1、3、6、12

    # 確認按鈕
    OKbtn=(By.XPATH,'/html/body/form/div[1]/div[7]/input[3]')
    # OKbtn=(By.CSS_SELECTOR,'input.cssButton')

    ### 契變類別(共用選項)
    EdorStyle=(By.ID,'EdorStyle')
    # EdorStyle=(By.XPATH,'//html/body/form/div[5]/div/table[1]/tbody/tr[1]/td[2]/input[1]')
    # EdorStyle=(By.ID,'EdorItemAppDate')

    # 繳別變更
    TypeChange=(By.XPATH,'//html/body/form/div[5]/div/table[2]/tbody/tr[2]/td[5]/input')

    # 補告知
    inform=(By.XPATH,'/html/body/form/div[3]/div/table[2]/tbody/tr[6]/td[2]/input')
    #EdorCodeListTable > tbody > tr:nth-child(6) > td:nth-child(2) > input[type=checkbox]
    # 補告知類型
    ImpartType=(By.ID,'ImpartType')

    # 增加契變按鈕
    AddItem=(By.ID,'AddItem')

    # 契變明細輸入按鈕(跳轉業面)
    DetailInput=(By.ID,'DetailInput')

    # 刪除契變項目
    DeleteItem=(By.ID,'DeleteItem')

    # 點擊序號
    EdorItemGridSel0=(By.ID,'EdorItemGridSel0')

    ### 變更後繳別(共用選項)
    PayIntv=(By.ID,'PayIntv')
    # PayIntv=(By.XPATH,'/html/body/form/div/div[4]/div/table/tbody/tr/td[2]/input[1]')

    # 保存
    saveBtn=(By.ID,'save')

    # 返回
    backBtn=(By.XPATH,'/html/body/form/div/div[2]/input[2]')

    # CRS按鈕
    sqButton11=(By.ID,'sqButton11')

    # 選擇CRS身分
    CrsIdentity=(By.ID,'CrsIdentity')

    # 選擇聲明日期
    statementDate=(By.ID,'statementDate')

    # 保存
    CRSSaveBtn=(By.PARTIAL_LINK_TEXT,'保存')

    # FACT按鈕(未完成)
    sqButton10=(By.ID,'sqButton10')

    # 是否檢附
    signin1=(By.XPATH,'/html/body/form/div[6]/table[1]/tbody/tr/td/span/div/div[3]/div/table/thead/tr/th[11]/div/input')
    # signin1=(By.CSS_SELECTOR,"input[type='checkbox']")

    # 要保人簽名
    signin2=(By.XPATH,'/html/body/form/div[6]/table[1]/tbody/tr/td/span/div/div[3]/div/table/thead/tr/th[12]/div/input')
    # 被保人簽名
    signin3=(By.XPATH,'/html/body/form/div[6]/table[1]/tbody/tr/td/span/div/div[3]/div/table/thead/tr/th[13]/div/input')

    # 法定代理人簽名
    signin4=(By.XPATH,'/html/body/form/div[6]/table[1]/tbody/tr/td/span/div/div[3]/div/table/thead/tr/th[14]/div/input')

    # 輸入完畢
    # submitBtn=(By.CSS_SELECTOR,'input.cssButton')
    submitBtn=(By.XPATH,'/html/body/form/div[14]/input')


# ----保全申請-----

    # 點擊保全申請
    def click_PreservationPage_ApplyBtn(self):
        self.find_element(*self.ApplyBtn).click()
        # 換視窗，裡面的1是變數可以決定更換為第幾個視窗
        self.switch_window(2)
        ### 確定跳轉視窗的名稱
        print(self.driver.title)
        self.switch_default()
        self.switch_frame("fraInterface")

    # 若已經申請過的就直接點水池內第一個選項
    def click_PreservationPage_PrivatePool(self,Policy_number):
        # 輸入欲查詢的保單號碼
        self.find_element(*self.PrivateWorkPoolQueryGrid3r0).send_keys(Policy_number)
        # 查詢按鈕
        self.find_element(*self.privateSearch).click()
        # 選擇池內第一個選項
        self.find_element(*self.PrivateWorkPoolGridSel0).click()
        # 點擊保全受理
        self.find_element(*self.PoolAccept).click()
        # 換視窗，裡面的1是變數可以決定更換為第幾個視窗
        self.switch_window(1)
        self.switch_default()
        self.switch_frame("fraInterface")

    # 輸入欲新增保全的保單號碼
    def input_PreservationPage_PolicyNumber(self,Policy_number):
        self.find_element(*self.input_PolicyNumber).send_keys(Policy_number)

    # 申請方式
    def select_PreservationPage_AppType(self):
        # 點擊申請方式
        self.find_element(*self.AppType).click()

        # 選擇申請方式
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('01')
        # 拍照
        self.take_screenshot()

        # 點擊確認
        # BtnOK=self.find_elements(*self.OKbtn)
        # BtnOK1=BtnOK[2]
        # ActionChains(self.driver).move_to_element(BtnOK1).click().perform()
        self.find_element(*self.OKbtn).click()

    # 選擇契變類別
    def select_PreservationPage_TypeChange(self):
        # 點擊契變類別
        # self.find_element(*self.EdorStyle).click()
        EdorStyle1=self.find_element(*self.EdorStyle)
        ActionChains(self.driver).move_to_element(EdorStyle1).double_click().perform()

        # 選擇契變類別 1一般契變
        c=self.find_element(*self.codeselect)
        Select(c).select_by_value('1')

        # 點擊補告知
        self.find_element(*self.inform).click()

        # 點擊增加契變項目按鈕
        self.find_element(*self.AddItem).click()

        # # 點擊序號前面圈圈
        # time.sleep(10)
        self.find_element(*self.EdorItemGridSel0).click()
        

        # 點擊契變明細輸入按鈕
        self.find_element(*self.DetailInput).click()

        # 跳轉至補告知頁面
        ### 跳轉視窗
        self.switch_window(4)
        ### 確定跳轉視窗的名稱
        print("測試:"+self.driver.title)
        self.switch_default()
        self.switch_frame("fraInterface")
        self.find_element(*self.ImpartType).send_keys('1')
        self.find_element(*self.ImpartType).send_keys(Keys.ENTER)

        # # 選擇保存
        self.find_element(*self.saveBtn).click()
        self.take_screenshot()

        # 選擇返回
        self.find_element(*self.backBtn).click()
        time.sleep(2)

    # 輸入CRS(會跳轉頁面)
    def FillIn_PreservationPage_CRS(self):
        # # 跳回原本窗口
        # self.switch_window_back()

        # 點擊CRS按鈕
        self.switch_window(1)
        self.switch_default()
        self.switch_frame("fraInterface")

        # CRS按鈕
        self.find_element(*self.sqButton11).click()

        # 跳轉至CRS頁面
        self.switch_window(3)
        self.switch_default()
        self.switch_frame("fraInterface")
        t0=self.driver.title
        print("目前窗口："+t0)

        # 點擊CRS身分選擇
        CrsIdentity1=self.find_element(*self.CrsIdentity)
        ActionChains(self.driver).move_to_element(CrsIdentity1).double_click().perform()

        # 選擇(不具外國稅務居民身分無須填寫自我證明(N則代表僅為台灣或同時具有台灣美國的稅務居民身分))
        d=self.find_element(*self.codeselect)
        Select(d).select_by_value('N')

        # 輸入聲明日期
        self.find_element(*self.statementDate).send_keys('2021-11-23')

        # 點擊保存
        self.find_element(*self.CRSSaveBtn).click()

    def signup_PreservationPage(self):
        # # 跳回原本窗口
        self.switch_window_back()
        self.switch_window(2)
        print("測試2:"+self.driver.title)
        self.switch_default()
        self.switch_frame("fraInterface")

        # js='document.querySelectorAll("input[type=checkbox]");'
        # abc=self.driver.execute_script(js)
        # abc[42].checked=True
        # abc[43].checked=True
        # abc[44].checked=True
        # abc[45].checked=True

        # CheckBtn=self.find_elements(*self.signin1)
        # CheckBtn42=CheckBtn[42]
        # ActionChains(self.driver).move_to_element(CheckBtn42).double_click().perform()

        self.find_element(*self.signin1).click()
        self.find_element(*self.signin2).click()
        self.find_element(*self.signin3).click()
        self.find_element(*self.signin4).click()

    def submit_PreservationPage(self):
        # submitBtn1=self.find_elements(*self.submitBtn)
        # submitBtn23=submitBtn1[23]
        # ActionChains(self.driver).move_to_element(submitBtn23).double_click().perform()
        self.find_element(*self.submitBtn).click()
        self.switch_window_back()


