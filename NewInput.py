from os import X_OK
from re import I
import time
from selenium.webdriver.common import by
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys


class NewInput():
    codeselect=(By.ID,'codeselect')
# ----------------------合同資訊-------------------------------
    # 保單號碼
    ProposalContNo=(By.ID,'ProposalContNo')

    # 管理機構
    ManageCom=(By.ID,'ManageCom')

    # 進件型態
    ProgressiveForm=(By.ID,'ProgressiveForm')

    # 電話要約/要保日期
    PolAppntDate=(By.ID,'PolAppntDate')

    # 生效日期
    CValiDate=(By.ID,'CValiDate')

    # 受理日期
    AcceptDate=(By.ID,'AcceptDate')

    # 要保書代碼
    PolicyCode=(By.ID,'PolicyCode')

    # 檢核文號
    ChecksNo=(By.ID,'ChecksNo')

    # OIU保單遞送方式
    OIUDelivery=(By.ID,'OIUDelivery')

    # 線上招攬
    OnlineAttract=(By.ID,'OnlineAttract')

    # 電子單據服務
    ElectronicDoc=(By.ID,'ElectronicDoc')

    # 錄音編號
    RecordNo=(By.ID,'RecordNo')

    # 是否速件
    WhetherQuickPolicy=(By.ID,'WhetherQuickPolicy')

    # 業務員代碼
    MultiAgentGridaddOne=(By.ID,'MultiAgentGridaddOne')

    MultiAgentGrid1r0=(By.ID,'MultiAgentGrid1r0')

    # 業務員姓名
    MultiAgentGrid2r0=(By.ID,'MultiAgentGrid2r0')
    # 業務員佣金比例
    MultiAgentGrid3r0=(By.ID,'MultiAgentGrid3r0')
    # 業務員代號/TM CODE
    MultiAgentGrid5r0=(By.ID,'MultiAgentGrid5r0')
    # TM CODE類別
    MultiAgentGrid6r0=(By.ID,'MultiAgentGrid6r0')

    # 舊客戶號碼
    IdNumber=(By.ID,'IdNumber')

    # 證件號碼
    AppntIDNo=(By.ID,'AppntIDNo')

    # 證件類型
    AppntIDType=(By.ID,'AppntIDType')

    # 證件有效期限至
    AppIDPeriodOfValidity=(By.ID,'AppIDPeriodOfValidity')

    # 姓名
    AppntName=(By.ID,'AppntName')
    # 英文姓名
    NameEn=(By.ID,'NameEn')

    # 與被保險人關係
    RelationToInsured=(By.ID,'RelationToInsured')
    # 原因
    Reason=(By.ID,'Reason')
    # 性別
    AppntSex=(By.ID,'AppntSex')
    # 出生日期
    AppntBirthday=(By.ID,'AppntBirthday')
    # 年齡
    AppntAge=(By.ID,'AppntAge')
    # 證件類型2 OthIDType
    OthIDType=(By.ID,'OthIDType')
    # 證件號碼2 OthIDNo
    OthIDNo=(By.ID,'OthIDNo')
    # 證件號碼2有效期至 OthIDExpDate
    OthIDExpDate=(By.ID,'OthIDExpDate')
    # 婚姻狀況
    AppntMarriage=(By.ID,'AppntMarriage')
    # 國籍
    AppntNativePlace=(By.ID,'AppntNativePlace')
    # 營業項目 BusinessProject
    BusinessProject=(By.ID,'BusinessProject')
    # 服務單位 ServiceUnit
    ServiceUnit=(By.ID,'ServiceUnit')
    # 職位 AppntPosition
    AppntPosition=(By.ID,'AppntPosition')
    # 職位代碼
    AppntOccupationCode=(By.ID,'AppntOccupationCode')
    # 職業等級
    AppntOccupationType=(By.ID,'AppntOccupationType')
    # 工作內容
    AppntWorkType=(By.ID,'AppntWorkType')

    # 行動電話 AppntMobile
    AppntMobile=(By.ID,'AppntMobile')
    # 聯絡電話（公司） AppntCompanyPhoneCode AppntCompanyPhone
    AppntCompanyPhoneCode=(By.ID,'AppntCompanyPhoneCode')
    AppntCompanyPhone=(By.ID,'AppntCompanyPhone')
    # 聯絡電話（住宅） AppntHomePhoneCode AppntHomePhone
    AppntHomePhoneCode=(By.ID,'AppntHomePhoneCode')
    AppntHomePhone=(By.ID,'AppntHomePhone')
    # 郵遞區號 AppntZipCode
    AppntZipCode=(By.ID,'AppntZipCode')
    # 住所 HomeCityName
    HomeCityName=(By.ID,'HomeCityName')
    HomeCountyName=(By.ID,'HomeCountyName')
    HomeDetailAddress=(By.ID,'HomeDetailAddress')
    # 住所地址 ForeignResidence
    ForeignResidence=(By.ID,'ForeignResidence')
    # E-mail AppntEMail
    AppntEMail=(By.ID,'AppntEMail')

    # 兼職代碼 AppntPartOccupationCode
    AppntPartOccupationCode=(By.ID,'AppntPartOccupationCode')

    # 兼業等級 AppntPartOccupationType
    AppntPartOccupationType=(By.ID,'AppntPartOccupationType')

    # 是否领有身心障礙手册或身心障礙証明 DisabilityFlag
    DisabilityFlag=(By.ID,'DisabilityFlag')
    # 是否受有監護宣告 CustodyFlag
    CustodyFlag=(By.ID,'CustodyFlag')
    # 保存按鈕
    addbutton=(By.ID,'addbutton')
    # 如要保人為被保險人本人，可免填本欄，請選擇
    SamePersonFlag=(By.ID,'SamePersonFlag')
    # 主/次被保險人資訊(保存按鈕) divAddInsuredButtun
    divAddInsuredButtun=(By.ID,'divAddInsuredButtun')
    # 傳真
    AppntFax=(By.ID,'AppntFax')

# ----------------------------------------------------------

    def NewInput_personinfo_input(self,personinfo):
        self.switch_default()
        self.switch_frame("fraInterface")

        # 保單號碼
        self.find_element(*self.ProposalContNo).click()
        self.take_screenshot_forInput('保單號碼')

        # 管理機構
        self.find_element(*self.ManageCom).click()
        self.take_screenshot_forInput('管理機構')

        # 進件型態
        self.find_element(*self.ProgressiveForm).click()
        self.take_screenshot_forInput('進件型態')
        self.find_element(*self.RecordNo).click()

        # 電話要約/要保日期 2021-11-19
        self.find_element(*self.PolAppntDate).send_keys(personinfo["PolAppntDate"])
        self.take_screenshot_forInput('電話要約-要保日期')
        

        # 生效日期 2021-11-19
        self.find_element(*self.CValiDate).send_keys(personinfo["PolAppntDate"])
        self.take_screenshot_forInput('生效日期')
        
        # 受理日期
        self.find_element(*self.AcceptDate).click()
        self.take_screenshot_forInput('受理日期')

        # 要保書代碼
        self.find_element(*self.PolicyCode).click()
        b=self.find_element(*self.codeselect)
        # personinfo["PolicyCode"] 2130
        Select(b).select_by_value(personinfo["PolicyCode"])
        # self.find_element(*self.RgtSourcesType).send_keys(Keys.ENTER)
        # self.find_element(*self.PolicyCode).send_keys("2130")
        self.take_screenshot_forInput('要保書代碼')

        # 檢核文號
        self.find_element(*self.ChecksNo).click()
        self.take_screenshot_forInput('檢核文號')

        # OIU保單遞送方式 1
        self.find_element(*self.OIUDelivery).click()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(personinfo["OIUDelivery"])
        self.take_screenshot_forInput('OIU保單遞送方式')

        # 線上招攬 N
        self.find_element(*self.OnlineAttract).click()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(personinfo["OnlineAttract"])
        self.take_screenshot_forInput('線上招攬')

        # 電子單據服務 N
        self.find_element(*self.ElectronicDoc).click()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(personinfo["ElectronicDoc"])
        self.take_screenshot_forInput('電子單據服務')

        # 錄音編號
        self.find_element(*self.RecordNo).send_keys("123")
        self.take_screenshot_forInput('錄音編號')

        # 是否速件
        self.find_element(*self.WhetherQuickPolicy).click()
        self.take_screenshot_forInput('是否速件')


        # 輸入業務員代碼
        self.find_element(*self.MultiAgentGridaddOne).click()
        self.find_element(*self.MultiAgentGrid1r0).send_keys(personinfo["Salesman"])
        self.find_element(*self.MultiAgentGrid1r0).click()
        Salesman=self.find_element(*self.MultiAgentGrid1r0)
        ActionChains(self.driver).move_to_element(Salesman).double_click().perform()
        # self.take_screenshot_forInput('業務員代碼')
        
        # 業務員姓名
        self.find_element(*self.MultiAgentGrid2r0).click()
        self.take_screenshot_forInput('業務員姓名')

        # 業務員代碼=業務員登錄證字號
        ActionChains(self.driver).move_to_element(Salesman).double_click().perform()
        self.take_screenshot_forInput('業務員登錄證字號')

        # 業務員佣金比例
        self.find_element(*self.MultiAgentGrid3r0).click()
        self.take_screenshot_forInput('業務員佣金比例')

        # 業務員代號/TM CODE
        self.find_element(*self.MultiAgentGrid5r0).click()
        self.take_screenshot_forInput('業務員代號-TM CODE')

        # TM CODE類別
        self.find_element(*self.MultiAgentGrid6r0).click()
        self.take_screenshot_forInput('TM-CODE類別')

        # 舊客戶證件號碼 IdNumber
        self.find_element(*self.IdNumber).send_keys(personinfo["AppntIDNo"])
        self.take_screenshot_forInput('舊戶證件號碼')
        self.find_element(*self.IdNumber).clear()

        # 滾動到住所位置
        self.scroll_To(*self.HomeCityName)


        # 證件類型
        self.find_element(*self.AppntIDType).click()
        self.take_screenshot_forInput('證件類型')


        # 證件號碼 personinfo["AppntIDNo"]
        self.find_element(*self.AppntIDNo).click()
        self.find_element(*self.AppntIDNo).send_keys(personinfo["AppntIDNo"])
        self.take_screenshot_forInput('證件號碼')

        # 證件有效期限至 AppIDPeriodOfValidity
        self.find_element(*self.AppIDPeriodOfValidity).send_keys(personinfo["PolAppntDate"])
        self.take_screenshot_forInput('證件有效期限至')

        # 姓名 personinfo["AppntName"]
        self.find_element(*self.AppntName).click()
        self.find_element(*self.AppntName).send_keys(personinfo["AppntName"])
        self.take_screenshot_forInput('姓名')

        # 英文姓名
        self.find_element(*self.NameEn).send_keys("test")
        self.take_screenshot_forInput('英文姓名')

        # 與被保險人關係personinfo["RelationToInsured"]
        self.find_element(*self.RelationToInsured).click()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(personinfo["RelationToInsured"])
        self.take_screenshot_forInput('與被保險人關係')

        # 原因Reason
        self.find_element(*self.Reason).send_keys("123456")
        self.take_screenshot_forInput('原因')

        # 性別AppntSex
        self.find_element(*self.AppntSex).click()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(personinfo["AppntSex"])
        self.take_screenshot_forInput('性別')

        # 出生日期 AppntBirthday
        self.find_element(*self.AppntBirthday).send_keys(personinfo["AppntBirthday"])
        self.take_screenshot_forInput('出生日期')

        # 保險年齡 AppntAge
        self.find_element(*self.AppntAge).click()
        self.take_screenshot_forInput('保險年齡')

        # 證件類型2 OthIDType
        self.find_element(*self.OthIDType).click()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('0')
        self.take_screenshot_forInput('證件類型2')

        # 證件號碼2 OthIDNo
        self.find_element(*self.OthIDNo).send_keys(personinfo["AppntIDNo"])
        self.take_screenshot_forInput('證件號碼2')

        # 證件號碼2有效期至 OthIDExpDate
        self.find_element(*self.OthIDExpDate).send_keys(personinfo["PolAppntDate"])
        self.take_screenshot_forInput('證件號碼2有效期至')

        # 婚姻狀況 AppntMarriage
        self.find_element(*self.AppntMarriage).click()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(personinfo["AppntMarriage"])
        self.take_screenshot_forInput('婚姻狀況')

        # 國籍 AppntNativePlace
        self.find_element(*self.AppntNativePlace).click()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value("TW")
        self.take_screenshot_forInput('國籍')

        # 營業項目 BusinessProject
        self.find_element(*self.BusinessProject).send_keys("1")
        self.take_screenshot_forInput('營業項目')

        # 服務單位 ServiceUnit
        self.find_element(*self.ServiceUnit).send_keys("1")
        self.take_screenshot_forInput('服務單位')

        # 職位 AppntPosition
        self.find_element(*self.AppntPosition).send_keys("1")
        self.take_screenshot_forInput('職位')

        # 工作內容 （含兼職）AppntWorkType
        self.find_element(*self.AppntWorkType).send_keys("123456")
        self.take_screenshot_forInput('工作內容含兼職')

        # 職位代碼 AppntOccupationCode
        self.find_element(*self.AppntOccupationCode).click()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(personinfo["AppntOccupationCode"])
        self.take_screenshot_forInput('職位代碼')

        # 職業等級 AppntOccupationType
        self.find_element(*self.AppntOccupationType).click()
        self.take_screenshot_forInput('職業等級')

        # 行動電話 AppntMobile
        self.find_element(*self.AppntMobile).send_keys('0987654321')
        self.take_screenshot_forInput('行動電話')

        # 聯絡電話（公司） AppntCompanyPhoneCode AppntCompanyPhone
        self.find_element(*self.AppntCompanyPhoneCode).send_keys('02')
        self.find_element(*self.AppntCompanyPhone).send_keys('12345678')
        self.take_screenshot_forInput('聯絡電話-公司')

        # 聯絡電話（住宅） AppntHomePhoneCode AppntHomePhone
        self.find_element(*self.AppntHomePhoneCode).send_keys('02')
        self.find_element(*self.AppntHomePhone).send_keys('12345678')
        self.take_screenshot_forInput('聯絡電話-住宅')

        # 郵遞區號 AppntZipCode
        self.find_element(*self.AppntZipCode).send_keys('110')
        self.take_screenshot_forInput('郵遞區號')

        # 住所 HomeCityName HomeCountyName
        # self.find_element(*self.HomeCityName).send_keys(personinfo["HomeCityName"])
        # self.find_element(*self.HomeCountyName).send_keys(personinfo["HomeCountyName"])
        js='document.getElementById("HomeDetailAddress").removeAttribute("readonly");'
        self.driver.execute_script(js)
        self.find_element(*self.HomeDetailAddress).send_keys("台北市信義區測試路55號")
        self.take_screenshot_forInput('住所')

        # 滾動到"如要保人為被保險人本人，可免填本欄"
        self.scroll_To(*self.DisabilityFlag)

        # 住所地址 ForeignResidence
        self.find_element(*self.ForeignResidence).send_keys('123456')
        self.take_screenshot_forInput('住所地址')

        # E-mail AppntEMail
        self.find_element(*self.AppntEMail).send_keys('test@gmail.com')
        self.take_screenshot_forInput('E-mail')

        # 兼職代碼 AppntPartOccupationCode
        self.find_element(*self.AppntPartOccupationCode).click()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(personinfo["AppntOccupationCode"])
        self.take_screenshot_forInput('兼職代碼')

        # 兼業等級 AppntPartOccupationType
        self.find_element(*self.AppntPartOccupationType).click()
        self.take_screenshot_forInput('兼業等級')

        # 是否领有身心障礙手册或身心障礙証明 DisabilityFlag
        self.find_element(*self.DisabilityFlag).send_keys(personinfo["DisabilityFlag"])
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(personinfo["DisabilityFlag"])
        self.find_element(*self.codeselect).click()
        self.take_screenshot_forInput('是否领有身心障礙手册或身心障礙証明')

        # 傳真
        self.find_element(*self.AppntFax).send_keys("123")
        self.take_screenshot_forInput('傳真')

        # 是否受有監護宣告 CustodyFlag
        # self.find_element(*self.AppntEMail).click()
        self.find_element(*self.CustodyFlag).send_keys(personinfo["CustodyFlag"])
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(personinfo["CustodyFlag"])
        self.find_element(*self.codeselect).click()
        self.take_screenshot_forInput('是否受有監護宣告')


        # 保存按鈕 addbutton
        self.find_element(*self.addbutton).click()
        time.sleep(3)
        # 如要保人為被保險人本人，可免填本欄，請選擇
        self.find_element(*self.SamePersonFlag).click()
        # 主/次被保險人資訊(保存按鈕)
        self.find_element(*self.divAddInsuredButtun).click()
        # 彈出alert視窗
        alert = self.driver.switch_to.alert
        # 列印彈窗顯示的資訊:Alert Message
        print(alert.text) 
        # 接受
        alert.accept()
        # -------------------------------------------------------------------------------------------------------------------------

   # ----------------------------------------------------------

    # 主/次被保險人資訊    證件號碼
    InsuredIdNo=(By.ID,'InsuredIdNo')
    # 主/次被保險人
    SequenceNo=(By.ID,'SequenceNo')
    # 與主被保險人關係
    RelationToMainInsured=(By.ID,'RelationToMainInsured')
    # 姓名
    Name=(By.ID,'Name')
    # 英文姓名
    InsuredNameEn=(By.ID,'InsuredNameEn')
    # 性別
    Sex=(By.ID,'Sex')
    # 出生日期
    Birthday=(By.ID,'Birthday')
    # 保險年齡
    InsuredAppAge=(By.ID,'InsuredAppAge')
    # 證件類型
    IDType=(By.ID,'IDType')
    # 證件號碼
    IDNo=(By.ID,'IDNo')
    # 證件號碼有效期至
    IDPeriodOfValidity=(By.ID,'IDPeriodOfValidity')
    # 證件類型2
    IDType2=(By.ID,'IDType2')
    # 證件號碼2
    IDNo2=(By.ID,'IDNo2')
    # 證件號碼2有效期至
    IDExpDate2=(By.ID,'IDExpDate2')
    # 婚姻狀況
    Marriage=(By.ID,'Marriage')
    # 國籍
    NativePlace=(By.ID,'NativePlace')
    # 營業項目
    InsuredBusinessProject=(By.ID,'InsuredBusinessProject')
    # 服務單位
    InsuredServiceUnit=(By.ID,'InsuredServiceUnit')
    # 職位
    InsuredPosition=(By.ID,'InsuredPosition')

    # 職業代碼
    OccupationCode=(By.ID,'OccupationCode')
    # 職業等級
    OccupationType=(By.ID,'OccupationType')
    # 工作內容含(兼職)
    WorkType=(By.ID,'WorkType')
    # 兼職代碼
    PartOccupationCode=(By.ID,'PartOccupationCode')
    # 兼業等級
    PartOccupationType=(By.ID,'PartOccupationType')
    # 是否參加全民健康保險
    SocialInsuFlag=(By.ID,'SocialInsuFlag')
    # 行動電話
    Mobile=(By.ID,'Mobile')
    # 聯絡電話（公司） CompanyPhoneCode CompanyPhone
    CompanyPhoneCode=(By.ID,'CompanyPhoneCode')
    CompanyPhone=(By.ID,'CompanyPhone')

    # 聯絡電話（住宅） HomePhoneCode  HomePhone
    HomePhoneCode=(By.ID,'HomePhoneCode')
    HomePhone=(By.ID,'HomePhone')
    # 郵遞區號 ZIPCODE
    ZIPCODE=(By.ID,'ZIPCODE')
    # 住所 PostalCityName
    PostalCityName=(By.ID,'PostalCityName')
    # 住所地址  InsuredForeignResidence
    InsuredForeignResidence=(By.ID,'InsuredForeignResidence')
    # 是否领有身心障礙手册或身心障礙証明  InsuredDisabilityFlag
    InsuredDisabilityFlag=(By.ID,'InsuredDisabilityFlag')
    # EMail
    EMail=(By.ID,'EMail')
    # 是否受有監護宣告 InsuredCustodyFlag
    InsuredCustodyFlag=(By.ID,'InsuredCustodyFlag')

    def NewInput_personinfo_input2(self,personinfo):
        self.switch_default()
        self.switch_frame("fraInterface")
                
        # # 如要保人為被保險人本人，可免填本欄，請選擇
        # self.find_element(*self.SamePersonFlag).click()

        self.scroll_To(*self.InsuredIdNo)

        # 主/次被保險人資訊    證件號碼(舊號碼)
        self.find_element(*self.InsuredIdNo).click()
        self.find_element(*self.InsuredIdNo).send_keys(personinfo["AppntIDNo"])
        self.take_screenshot_forInput('舊客戶資料-證件號碼x')
        self.find_element(*self.InsuredIdNo).clear()

        # 主/次被保險人
        self.find_element(*self.SequenceNo).click()
        self.take_screenshot_forInput('主-次被保險人x')

        # 與主被保險人關係
        self.find_element(*self.RelationToMainInsured).click()
        self.take_screenshot_forInput('與主被保險人關係x')

        # 姓名 personinfo["AppntName"]
        self.find_element(*self.Name).click()
        self.take_screenshot_forInput('姓名x')

        # 英文姓名
        self.find_element(*self.InsuredNameEn).click()
        self.take_screenshot_forInput('英文姓名x')

        # 性別AppntSex
        self.find_element(*self.Sex).click()
        self.take_screenshot_forInput('性別x')

        # 出生日期 AppntBirthday
        self.find_element(*self.Birthday).click()
        self.take_screenshot_forInput('出生日期x')

        # 保險年齡 AppntAge
        self.find_element(*self.InsuredAppAge).click()
        self.take_screenshot_forInput('保險年齡x')

        # 證件類型
        self.find_element(*self.IDType).click()
        self.take_screenshot_forInput('證件類型x')

        # # 證件號碼 personinfo["AppntIDNo"]
        self.find_element(*self.IDNo).click()
        self.take_screenshot_forInput('證件號碼x')

        # 證件有效期限至 AppIDPeriodOfValidity
        self.find_element(*self.IDPeriodOfValidity).clear()
        self.find_element(*self.IDPeriodOfValidity).send_keys(personinfo["PolAppntDate"])
        self.take_screenshot_forInput('證件有效期限至x')

        # 證件類型2 OthIDType
        self.find_element(*self.IDType2).click()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('0')
        self.take_screenshot_forInput('證件類型2x')

        # 證件號碼2 OthIDNo
        self.find_element(*self.IDNo2).click()
        self.take_screenshot_forInput('證件號碼2x')

        # 證件號碼2有效期至 OthIDExpDate
        self.find_element(*self.IDExpDate2).clear()
        self.find_element(*self.IDExpDate2).send_keys(personinfo["PolAppntDate"])
        self.take_screenshot_forInput('證件號碼2有效期至x')

        # # 婚姻狀況 AppntMarriage
        self.find_element(*self.Marriage).click()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(personinfo["AppntMarriage"])
        self.take_screenshot_forInput('婚姻狀況x')

        # 國籍 AppntNativePlace
        self.find_element(*self.NativePlace).click()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value("TW")
        self.take_screenshot_forInput('國籍x')
        self.find_element(*self.InsuredBusinessProject).click()

         # 營業項目 BusinessProject
        self.find_element(*self.InsuredBusinessProject).clear()
        self.find_element(*self.InsuredBusinessProject).send_keys("1")
        self.take_screenshot_forInput('營業項目x')

        
        # 服務單位 ServiceUnit
        self.find_element(*self.InsuredServiceUnit).clear()
        self.find_element(*self.InsuredServiceUnit).send_keys("1")
        self.take_screenshot_forInput('服務單位x')

        # 職位 AppntPosition
        self.find_element(*self.InsuredPosition).clear()
        self.find_element(*self.InsuredPosition).send_keys("1")
        self.take_screenshot_forInput('職位x')

        # 工作內容 （含兼職）AppntWorkType
        self.find_element(*self.WorkType).clear()
        self.find_element(*self.WorkType).send_keys("123456")
        self.take_screenshot_forInput('工作內容-含兼職x')

        # 職位代碼 AppntOccupationCode
        self.find_element(*self.OccupationCode).click()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(personinfo["AppntOccupationCode"])
        self.take_screenshot_forInput('職位代碼x')

        # # 職業等級 AppntOccupationType
        self.find_element(*self.OccupationType).click()
        self.take_screenshot_forInput('職業等級x')

        # 行動電話 Mobile
        self.find_element(*self.Mobile).click()
        self.take_screenshot_forInput('行動電話x')

        # 聯絡電話（公司） CompanyPhoneCode CompanyPhone
        self.find_element(*self.CompanyPhoneCode).click()
        self.find_element(*self.CompanyPhone).click()
        self.take_screenshot_forInput('聯絡電話-公司x')


        # 聯絡電話（住宅） HomePhoneCode  HomePhone
        self.find_element(*self.HomePhoneCode).click()
        self.find_element(*self.HomePhone).click()
        self.take_screenshot_forInput('聯絡電話-住宅x')

        # # 滾動到"是否受有監護宣告"
        self.scroll_To(*self.InsuredCustodyFlag)

        # 郵遞區號 ZIPCODE
        self.find_element(*self.ZIPCODE).click()
        self.take_screenshot_forInput('郵遞區號x')

        # 住所 PostalCityName
        self.find_element(*self.PostalCityName).click()
        self.take_screenshot_forInput('住所x')

        # #住所地址  InsuredForeignResidence
        self.find_element(*self.InsuredForeignResidence).click()
        self.take_screenshot_forInput('住所地址x')

        
        # 是否參加全民健康保險
        self.find_element(*self.SocialInsuFlag).click()
        self.take_screenshot_forInput('是否參加全民健康保險x')

        # 兼職代碼 AppntPartOccupationCode
        self.find_element(*self.PartOccupationCode).click()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(personinfo["AppntOccupationCode"])
        self.take_screenshot_forInput('兼職代碼x')

        # 兼業等級 AppntPartOccupationType
        self.find_element(*self.PartOccupationType).click()
        self.take_screenshot_forInput('兼業等級x')

        # 是否领有身心障礙手册或身心障礙証明  InsuredDisabilityFlag
        # # 是否领有身心障礙手册或身心障礙証明 DisabilityFlag
        self.find_element(*self.InsuredDisabilityFlag).click()
        self.take_screenshot_forInput('是否领有身心障礙手册或身心障礙証明x')

        ## EMail
        self.find_element(*self.EMail).click()
        self.take_screenshot_forInput('EMailx')

        # 是否受有監護宣告 InsuredCustodyFlag
        self.find_element(*self.InsuredCustodyFlag).click()
        self.take_screenshot_forInput('是否受有監護宣告x')


        # # 傳真
        # self.find_element(*self.AppntFax).send_keys("123")
        # self.take_screenshot_forInput()

    # 商品錄入按鈕
    Productbtn=(By.XPATH,'/html/body/form/input[1]')
    # 進入商品資訊錄入畫面
    back=(By.ID,'back')
    # 刪除按鈕
    deleBtn=(By.ID,'goodDeleteButt')
    # 選擇商品
    PolGridSel0=(By.ID,'PolGridSel0')
    # 保存按鈕
    Savebtn=(By.XPATH,'/html/body/form/div[30]/div/div[1]/table/tbody/tr/td[2]/input')
    # 添加按鈕
    addBtn=(By.XPATH,'/html/body/form/div[30]/div/div[5]/div/input[1]')
    # 批註條款_添加按鈕
    PostilGridaddOne=(By.ID,'PostilGridaddOne')
    # 上一步Btn
    backBtn=(By.ID,'divInputContButton')
    # backBtn=(By.XPATH,'/html/body/form/div[30]/div/div[9]/input[3]')

    # 險種編碼
    RiskCode=(By.ID,'RiskCode')
    # 保額
    Amnt=(By.ID,'Amnt')
    # 型別
    # PolType=(By.XPATH,'/html/body/form/div[16]/table/tbody/tr[1]/td[6]/input[1]')
    PolType=(By.CSS_SELECTOR,'input[name="PolType"]')

    # 第一期保費責任(保費) //html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[4]
    priceCol=(By.XPATH,'//html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[4]')
    priceBtn=(By.CSS_SELECTOR,'input.datagrid-editable-input')


    # 投資標的受益分配之約定帳戶-帳號 
    GetBankAccNo=(By.ID,'GetBankAccNo')
    # 投資標的受益分配之約定帳戶-戶名 GetAccName
    GetAccName=(By.ID,'GetAccName')
    # 投資標的受益分配之約定帳戶-銀行與分行 GetBankCode
    GetBankCode=(By.ID,'GetBankCode')
    # 投資標的代碼 InvestPlanRate1r0
    investCode=(By.ID,'InvestPlanRate1r0')
    # 投資趴數
    InvestPlanRate5r0=(By.ID,'InvestPlanRate5r0')
    # 每月扣除順序 InvestPlanRate7r1
    InvestPlanRate7r0=(By.ID,'InvestPlanRate7r0')

    # 保費 Prem
    Prem=(By.ID,'Prem')
    # 幣別
    CurrencyCode=(By.ID,'CurrencyCode')
    # 繳別 PayIntv
    PayIntv=(By.ID,'PayIntv')
    # 繳費期間
    PayEndYear=(By.ID,'PayEndYear')
    # 繳費期間單位 
    PayEndYearFlag=(By.ID,'PayEndYearFlag')
    # 保證領取期間
    Guaryear=(By.ID,'Guaryear')
    # 保證領取期間單位
    GuaryearFlag=(By.ID,'GuaryearFlag')

    # 資產撥回給付方式 /html/body/form/div[16]/table/tbody/tr[4]/td[4]/input[1]
    assectWay=(By.XPATH,'/html/body/form/div[16]/table/tbody/tr[4]/td[4]/input[1]')
    assectWayNIFU0101=(By.CSS_SELECTOR,'input[name="AssetBackGetMode"]')
    # 增值回饋分享金給付方式
    BonusGetMode=(By.ID,'BonusGetMode')
    # 自動墊繳
    AutoPayFlag=(By.XPATH,'/html/body/form/div[16]/table/tbody/tr[4]/td[4]/input[1]')

    # 批註條款代號
    PostilGrid1r0=(By.ID,'PostilGrid1r0')

    # 原投資代號 PostilConsultGrid1r0
    PostilConsultGrid1r0=(By.ID,'PostilConsultGrid1r0')
    # 原投資代號名稱 PostilConsultGrid2r0
    PostilConsultGrid2r0=(By.ID,'PostilConsultGrid2r0')
    # 新投資代號 PostilConsultGrid3r0
    PostilConsultGrid3r0=(By.ID,'PostilConsultGrid3r0')
    # 新投資代號名稱 PostilConsultGrid4r0
    PostilConsultGrid4r0=(By.ID,'PostilConsultGrid4r0')

    # 年金給付方式
    Payway=(By.XPATH,'/html/body/form/div[16]/table/tbody/tr[2]/td[2]/input[1]')
    PaywayNITA0901=(By.XPATH,'/html/body/form/div[16]/table/tbody/tr[3]/td[6]/input[1]')
    NewPayway=(By.CSS_SELECTOR,'input[name="GetIntv"]')

    # 年金給付開始年齡 GetYear
    GetYear=(By.ID,'GetYear')

    # 領取方式之選擇 /html/body/form/div[16]/table/tbody/tr[2]/td[4]/input[1]
    Payselect=(By.XPATH,'/html/body/form/div[16]/table/tbody/tr[2]/td[4]/input[1]')

    # 領取之開始日 GuarStartDate
    GuarStartDate=(By.ID,'GuarStartDate')

    # 領取之開始年齡 GuarStartYear
    GuarStartYear=(By.ID,'GuarStartYear')

    # 增值回饋分享金給付帳戶-銀行
    BonusGetBankCode=(By.ID,'BonusGetBankCode')
    # 帳號
    BonusGetBankAccNo=(By.ID,'BonusGetBankAccNo')
    # 戶名
    BonusGetAccName=(By.ID,'BonusGetAccName')


    def NewInput_Product_input(self,product):
        self.switch_default()
        self.switch_frame("fraInterface")

     # --商品錄入------------------------------------------------------
        # 商品錄入按鈕
        self.find_element(*self.Productbtn).click()

        # 險種編碼 NTIW0701 
        RiskCode2=self.find_element(*self.RiskCode)
        ActionChains(self.driver).move_to_element(RiskCode2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(product["productId"])
        self.take_screenshot_forInput('險種編碼')

        # 進入商品資訊錄入畫面
        self.find_element(*self.back).click()

        # 保額
        self.find_element(*self.Amnt).send_keys("5000000")
        self.take_screenshot_forInput('保額')

        # 保費 Prem
        self.find_element(*self.Prem).click()
        self.take_screenshot_forInput('保費')

        # 幣別
        self.find_element(*self.CurrencyCode).click()
        self.take_screenshot_forInput('幣別')
        

        # 繳別 PayIntv
        PayIntv2=self.find_element(*self.PayIntv)
        ActionChains(self.driver).move_to_element(PayIntv2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(product["PayIntv"])
        self.take_screenshot_forInput('繳別')

        # 繳費期間
        PayEndYear2=self.find_element(*self.PayEndYear)
        ActionChains(self.driver).move_to_element(PayEndYear2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(product["PayEndYear"])
        self.take_screenshot_forInput('繳費期間')

        # 增值回饋分享金給付方式
        BonusGetMode2=self.find_element(*self.BonusGetMode)
        ActionChains(self.driver).move_to_element(BonusGetMode2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('1')
        self.take_screenshot_forInput('增值回饋分享金給付方式')

        # 自動墊繳
        AutoPayFlag2=self.find_element(*self.AutoPayFlag)
        ActionChains(self.driver).move_to_element(AutoPayFlag2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('0')
        self.take_screenshot_forInput('自動墊繳')

        # 保存
        self.find_element(*self.Savebtn).click()

        # 增值回饋分享金給付帳戶
        BonusGetBankCode2=self.find_element(*self.BonusGetBankCode)
        ActionChains(self.driver).move_to_element(BonusGetBankCode2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('8220015')
        self.take_screenshot_forInput('增值回饋分享金給付帳戶-銀行')

        # 帳號
        self.find_element(*self.BonusGetBankAccNo).click()
        self.find_element(*self.BonusGetBankAccNo).send_keys('0123456789')
        self.take_screenshot_forInput('增值回饋分享金給付帳戶-帳號')
        
        # 戶名
        self.find_element(*self.BonusGetAccName).click()
        self.take_screenshot_forInput('增值回饋分享金給付帳戶-戶名')

        # 移動
        self.scroll_To(*self.PostilGrid1r0)

        # # 強制執行上一步
        self.driver.execute_script("backBtn = document.querySelector('#riskbutton1');" + "backBtn.onclick();")

        # 選擇商品
        self.find_element(*self.PolGridSel0).click()
        
        # 刪除按鈕
        self.find_element(*self.deleBtn).click()

     # --年金給付開始年齡------------------------------------------------------
        # 商品錄入按鈕
        self.find_element(*self.Productbtn).click()

        # 險種編碼  NITA0901
        RiskCode2=self.find_element(*self.RiskCode)
        ActionChains(self.driver).move_to_element(RiskCode2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('NIFA0102')

        # 進入商品資訊錄入畫面
        self.find_element(*self.back).click()

        # 保費
        self.find_element(*self.Prem).send_keys('5000000')

        # 年金給付方式/html/body/form/div[16]/table/tbody/tr[2]/td[2]/input[1]
        Payway2=self.find_element(*self.Payway)
        ActionChains(self.driver).move_to_element(Payway2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('0')
        self.take_screenshot_forInput('年金給付方式')

        # 年金給付開始年齡 GetYear
        self.find_element(*self.GetYear).send_keys('50')
        self.take_screenshot_forInput('年金給付開始年齡')

        # 領取方式之選擇 /html/body/form/div[16]/table/tbody/tr[2]/td[4]/input[1]
        Payselect2=self.find_element(*self.Payselect)
        ActionChains(self.driver).move_to_element(Payselect2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('1')
        self.take_screenshot_forInput('領取方式之選擇')

        # 領取之開始日 GuarStartDate
        self.find_element(*self.GuarStartDate).click()
        self.find_element(*self.GuarStartDate).send_keys('2021-11-19')
        self.take_screenshot_forInput('領取之開始日')

        # 領取之開始年齡 GuarStartYear
        self.find_element(*self.GuarStartYear).click()
        self.find_element(*self.GuarStartYear).send_keys('50')
        self.take_screenshot_forInput('領取之開始年齡')

        # # 強制執行上一步
        self.driver.execute_script("backBtn = document.querySelector('#riskbutton1');" + "backBtn.onclick();")



     # --投資標配置及比例------------------------------------------------------

        # 商品錄入按鈕
        self.find_element(*self.Productbtn).click()

        # 險種編碼 NITU0401 = NIFU0501
        RiskCode2=self.find_element(*self.RiskCode)
        ActionChains(self.driver).move_to_element(RiskCode2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value("NITU0401")

        # 進入商品資訊錄入畫面
        self.find_element(*self.back).click()

        # 保額
        self.find_element(*self.Amnt).send_keys("5000000")

        # 型別
        PolType2=self.find_element(*self.PolType)
        ActionChains(self.driver).move_to_element(PolType2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value("1")
        self.take_screenshot_forInput('型別')

        # 第一期保費責任(點擊外圍框框) //html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[4]
        priceCol2=self.find_element(*self.priceCol)
        ActionChains(self.driver).move_to_element(priceCol2).double_click().perform()
        # 第一期保費責任(保費)
        ele=self.find_elements(*self.priceBtn)
        ele2=ele[2]
        ele2.send_keys("1234")

        # 保存
        self.find_element(*self.Savebtn).click()

        # -----------投資標的收益分配或資產撥回之約定帳戶------------------
        # 投資標的受益分配之約定帳戶-帳號 GetBankAccNo
        self.find_element(*self.GetBankAccNo).click()
        self.find_element(*self.GetBankAccNo).send_keys('0123456789')
        self.take_screenshot_forInput('投資標的受益分配之約定帳戶-帳號')

        # 投資標的受益分配之約定帳戶-戶名 GetAccName
        self.find_element(*self.GetAccName).click()
        self.take_screenshot_forInput('投資標的受益分配之約定帳戶-戶名')

        # 投資標的受益分配之約定帳戶-銀行與分行 GetBankCode
        self.find_element(*self.GetBankCode).click()
        self.find_element(*self.GetBankCode).send_keys('8220015')
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('8220015')
        self.find_element(*self.codeselect).click()
        self.take_screenshot_forInput('投資標的受益分配之約定帳戶-銀行與分行')

        # 投資標的代碼 InvestPlanRate1r0
        self.find_element(*self.investCode).send_keys(product["investTarget2"])
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(product["investTarget2"])
        self.find_element(*self.codeselect).click()
        self.take_screenshot_forInput('投資標的代碼')

        # 投資趴數
        self.find_element(*self.InvestPlanRate5r0).click()
        self.find_element(*self.InvestPlanRate5r0).send_keys("100")
        self.take_screenshot_forInput('投資趴數')

        # 每月扣除順序
        self.find_element(*self.InvestPlanRate7r0).click()
        self.find_element(*self.InvestPlanRate7r0).send_keys('1')
        self.take_screenshot_forInput('每月扣除順序')

        # 添加按鈕
        self.find_element(*self.addBtn).click()

        # 移動 添加按鈕
        self.scroll_To(*self.addBtn)

        # # 強制執行上一步
        self.driver.execute_script("backBtn = document.querySelector('#riskbutton1');" + "backBtn.onclick();")


        # 選擇商品
        self.find_element(*self.PolGridSel0).click()
        # 刪除按鈕
        self.find_element(*self.deleBtn).click()
        # 商品錄入按鈕
        self.find_element(*self.Productbtn).click()
        # 險種編碼 NTIW0701
        RiskCode2=self.find_element(*self.RiskCode)
        ActionChains(self.driver).move_to_element(RiskCode2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(product["productId"])
        # 進入商品資訊錄入畫面
        self.find_element(*self.back).click()
        # 保額
        self.find_element(*self.Amnt).send_keys("5000000")
        # 繳別 PayIntv
        PayIntv2=self.find_element(*self.PayIntv)
        ActionChains(self.driver).move_to_element(PayIntv2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(product["PayIntv"])
        # 繳費期間
        PayEndYear2=self.find_element(*self.PayEndYear)
        ActionChains(self.driver).move_to_element(PayEndYear2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(product["PayEndYear"])
        # 增值回饋分享金給付方式
        BonusGetMode2=self.find_element(*self.BonusGetMode)
        ActionChains(self.driver).move_to_element(BonusGetMode2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('1')
        # 自動墊繳
        AutoPayFlag2=self.find_element(*self.AutoPayFlag)
        ActionChains(self.driver).move_to_element(AutoPayFlag2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('0')

        # 保存
        self.find_element(*self.Savebtn).click()

        # 增值回饋分享金給付帳戶
        BonusGetBankCode2=self.find_element(*self.BonusGetBankCode)
        ActionChains(self.driver).move_to_element(BonusGetBankCode2).double_click().perform()
        b=self.find_element(*self.codeselect)

    # 1.險種編碼 NITU0401 = NIFU0501(台灣人壽愛豐收外幣變額萬能壽險) 9000150778
    def product_NIFU0501(self,product):
        self.switch_default()
        self.switch_frame("fraInterface")
        # 商品錄入按鈕
        self.find_element(*self.Productbtn).click()

        # 險種編碼
        RiskCode2=self.find_element(*self.RiskCode)
        ActionChains(self.driver).move_to_element(RiskCode2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(product["productId"])
        self.take_screenshot_forInput('險種編碼')

        # 進入商品資訊錄入畫面
        self.find_element(*self.back).click()

        # 保額
        self.find_element(*self.Amnt).send_keys("5000000")
        self.take_screenshot_forInput('保額')

        # 型別
        PolType2=self.find_element(*self.PolType)
        ActionChains(self.driver).move_to_element(PolType2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value("1")
        self.take_screenshot_forInput('型別')

        # 第一期保費責任(點擊外圍框框) //html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[4]
        priceCol2=self.find_element(*self.priceCol)
        ActionChains(self.driver).move_to_element(priceCol2).double_click().perform()
        # 第一期保費責任(保費)
        ele=self.find_elements(*self.priceBtn)
        ele2=ele[2]
        ele2.send_keys("1234")
        self.take_screenshot_forInput('保費')

        # 保存
        self.find_element(*self.Savebtn).click()

        # -----------投資標的收益分配或資產撥回之約定帳戶------------------
        # 投資標的受益分配之約定帳戶-帳號 GetBankAccNo
        self.find_element(*self.GetBankAccNo).click()
        self.find_element(*self.GetBankAccNo).send_keys('0123456789')
        self.take_screenshot_forInput('投資標的受益分配之約定帳戶-帳號')

        # 投資標的受益分配之約定帳戶-戶名 GetAccName
        self.find_element(*self.GetAccName).click()
        self.take_screenshot_forInput('投資標的受益分配之約定帳戶-戶名')

        # 投資標的受益分配之約定帳戶-銀行與分行 GetBankCode
        self.find_element(*self.GetBankCode).click()
        self.find_element(*self.GetBankCode).send_keys('8220015')
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('8220015')
        self.find_element(*self.codeselect).click()
        self.take_screenshot_forInput('投資標的受益分配之約定帳戶-銀行與分行')

        # 投資標的代碼 InvestPlanRate1r0
        self.find_element(*self.investCode).send_keys(product["investTarget2"])
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(product["investTarget2"])
        self.find_element(*self.codeselect).click()
        self.take_screenshot_forInput('投資標的代碼')

        # 投資趴數
        self.find_element(*self.InvestPlanRate5r0).click()
        self.find_element(*self.InvestPlanRate5r0).send_keys("100")
        self.take_screenshot_forInput('投資趴數')

        # 每月扣除順序
        self.find_element(*self.InvestPlanRate7r0).click()
        self.find_element(*self.InvestPlanRate7r0).send_keys('1')
        self.take_screenshot_forInput('每月扣除順序')

        # 添加按鈕
        self.find_element(*self.addBtn).click()

        # 移動 添加按鈕
        self.scroll_To(*self.addBtn)

        # # 強制執行上一步
        self.driver.execute_script("backBtn = document.querySelector('#riskbutton1');" + "backBtn.onclick();")

    # 2.險種編碼 NIFA0102 = NITA0901(台灣人壽金采100變額年金保險) 9000150779
    def product_NITA0901(self,product):
        self.switch_default()
        self.switch_frame("fraInterface")
        # 商品錄入按鈕
        self.find_element(*self.Productbtn).click()

        # 險種編碼
        RiskCode2=self.find_element(*self.RiskCode)
        ActionChains(self.driver).move_to_element(RiskCode2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(product["productId"])
        self.take_screenshot_forInput('險種編碼')

        # 進入商品資訊錄入畫面
        self.find_element(*self.back).click()

        # 保費
        self.find_element(*self.Prem).send_keys('5000000')
        self.take_screenshot_forInput('保費')


        # 繳別 PayIntv
        self.find_element(*self.PayIntv).click()
        self.take_screenshot_forInput('繳別')

        # 繳費期間
        self.find_element(*self.PayEndYear).click()
        self.take_screenshot_forInput('繳費期間')

        # 繳費期間單位 PayEndYearFlag
        self.find_element(*self.PayEndYearFlag).click()
        self.take_screenshot_forInput('繳費期間單位')

        # 保證領取期間Guaryear
        self.find_element(*self.Guaryear).click()
        self.take_screenshot_forInput('保證領取期間')

        # 保證領取期間單位GuaryearFlag
        self.find_element(*self.GuaryearFlag).click()
        self.take_screenshot_forInput('保證領取期間單位')

        # 年金給付方式/html/body/form/div[16]/table/tbody/tr[2]/td[2]/input[1]
        Payway2=self.find_element(*self.PaywayNITA0901)
        ActionChains(self.driver).move_to_element(Payway2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('12')
        self.take_screenshot_forInput('年金給付方式')

        # 年金給付開始年齡 GetYear
        self.find_element(*self.GetYear).send_keys('50')
        self.take_screenshot_forInput('年金給付開始年齡')

        # 資產撥回給付方式 /html/body/form/div[16]/table/tbody/tr[4]/td[4]/input[1]
        assectWay2=self.find_element(*self.assectWay)
        ActionChains(self.driver).move_to_element(assectWay2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('0')
        self.take_screenshot_forInput('資產撥回給付方式')
        

        # 保存
        self.find_element(*self.Savebtn).click()

# -----------投資標的收益分配或資產撥回之約定帳戶------------------
        # 投資標的受益分配之約定帳戶-帳號 GetBankAccNo
        self.find_element(*self.GetBankAccNo).click()
        self.find_element(*self.GetBankAccNo).send_keys('0123456789')
        self.take_screenshot_forInput('投資標的受益分配之約定帳戶-帳號')

        # 投資標的受益分配之約定帳戶-戶名 GetAccName
        self.find_element(*self.GetAccName).click()
        self.take_screenshot_forInput('投資標的受益分配之約定帳戶-戶名')

        # 投資標的受益分配之約定帳戶-銀行與分行 GetBankCode
        self.find_element(*self.GetBankCode).click()
        self.find_element(*self.GetBankCode).send_keys('8220015')
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('8220015')
        self.find_element(*self.codeselect).click()
        self.take_screenshot_forInput('投資標的受益分配之約定帳戶-銀行與分行')

        # 投資標的代碼 InvestPlanRate1r0
        self.find_element(*self.investCode).send_keys(product["investTarget2"])
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(product["investTarget2"])
        self.find_element(*self.codeselect).click()
        self.take_screenshot_forInput('投資標的代碼')

        # 投資趴數
        self.find_element(*self.InvestPlanRate5r0).click()
        self.find_element(*self.InvestPlanRate5r0).send_keys("100")
        self.take_screenshot_forInput('投資趴數')

        # 每月扣除順序
        self.find_element(*self.InvestPlanRate7r0).click()
        self.find_element(*self.InvestPlanRate7r0).send_keys('1')
        self.take_screenshot_forInput('每月扣除順序')

        # 添加按鈕
        self.find_element(*self.addBtn).click()

        # 移動 添加按鈕
        self.scroll_To(*self.addBtn)

        # 批註條款_添加按鈕
        self.find_element(*self.PostilGridaddOne).click()

        # 條款代號 NIEO1101 PostilGrid1r0
        PostilGrid1r0x=self.find_element(*self.PostilGrid1r0)
        ActionChains(self.driver).move_to_element(PostilGrid1r0x).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('NIEO1101')

        # 原投資代號 PostilConsultGrid1r0
        self.find_element(*self.PostilConsultGrid1r0).click()
        self.take_screenshot_forInput('原投資代號')

        # 原投資代號名稱 PostilConsultGrid2r0
        self.find_element(*self.PostilConsultGrid2r0).click()
        self.take_screenshot_forInput('原投資代號名稱')
        
        # 新投資代號 PostilConsultGrid3r0
        PostilConsultGrid3r0x=self.find_element(*self.PostilConsultGrid3r0)
        ActionChains(self.driver).move_to_element(PostilConsultGrid3r0x).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('E077')
        self.take_screenshot_forInput('新投資代號')

        # 新投資代號名稱 PostilConsultGrid4r0
        self.find_element(*self.PostilConsultGrid4r0).click()
        self.take_screenshot_forInput('新投資代號名稱')


        # # 強制執行上一步
        self.driver.execute_script("backBtn = document.querySelector('#riskbutton1');" + "backBtn.onclick();")

    def product_NIFU0101(self,product):
        self.switch_default()
        self.switch_frame("fraInterface")
        # 商品錄入按鈕
        self.find_element(*self.Productbtn).click()

        # 險種編碼
        RiskCode2=self.find_element(*self.RiskCode)
        ActionChains(self.driver).move_to_element(RiskCode2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(product["productId"])
        self.take_screenshot_forInput('險種編碼')

        # 進入商品資訊錄入畫面
        self.find_element(*self.back).click()

        # 保費
        self.find_element(*self.Prem).click()
        self.find_element(*self.Prem).send_keys("10000")
        self.take_screenshot_forInput('保費')

        # 保額
        self.find_element(*self.Amnt).click()
        self.find_element(*self.Amnt).send_keys("5000000")
        self.take_screenshot_forInput('保額')

        # 幣別
        CurrencyCode2=self.find_element(*self.CurrencyCode)
        ActionChains(self.driver).move_to_element(CurrencyCode2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value("USD")
        self.take_screenshot_forInput('幣別')


        # 繳別 PayIntv
        self.find_element(*self.PayIntv).click()
        self.take_screenshot_forInput('繳別')

        # 繳費期間
        self.find_element(*self.PayEndYear).click()
        self.take_screenshot_forInput('繳費期間')

        # 繳費期間單位 PayEndYearFlag
        self.find_element(*self.PayEndYearFlag).click()
        self.take_screenshot_forInput('繳費期間單位')

        # 型別
        PolType2=self.find_element(*self.PolType)
        ActionChains(self.driver).move_to_element(PolType2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value("1")
        self.take_screenshot_forInput('型別')

        # 資產撥回給付方式  /html/body/form/div[16]/table/tbody/tr[4]/td[4]/input[1]
        assectWay2=self.find_element(*self.assectWay)
        ActionChains(self.driver).move_to_element(assectWay2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('0')
        self.take_screenshot_forInput('資產撥回給付方式')

        # 保存
        self.find_element(*self.Savebtn).click()

        # 投資標的代碼 InvestPlanRate1r0
        self.find_element(*self.investCode).send_keys(product["investTarget2"])
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(product["investTarget2"])
        self.find_element(*self.codeselect).click()
        self.take_screenshot_forInput('投資標的代碼')

        # 投資趴數
        self.find_element(*self.InvestPlanRate5r0).click()
        self.find_element(*self.InvestPlanRate5r0).send_keys("100")
        self.take_screenshot_forInput('投資趴數')

        # 每月扣除順序
        self.find_element(*self.InvestPlanRate7r0).click()
        self.find_element(*self.InvestPlanRate7r0).send_keys('1')
        self.take_screenshot_forInput('每月扣除順序')

        # 添加按鈕
        self.find_element(*self.addBtn).click()

        # 投資標的受益分配之約定帳戶-帳號 GetBankAccNo
        self.find_element(*self.GetBankAccNo).click()
        self.find_element(*self.GetBankAccNo).send_keys('0123456789')
        self.take_screenshot_forInput('投資標的受益分配之約定帳戶-帳號')

        # 投資標的受益分配之約定帳戶-戶名 GetAccName
        self.find_element(*self.GetAccName).click()
        self.take_screenshot_forInput('投資標的受益分配之約定帳戶-戶名')

        # 投資標的受益分配之約定帳戶-銀行與分行 GetBankCode
        self.find_element(*self.GetBankCode).click()
        self.find_element(*self.GetBankCode).send_keys('8220015')
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('8220015')
        self.find_element(*self.codeselect).click()
        self.take_screenshot_forInput('投資標的受益分配之約定帳戶-銀行與分行')

        # 移動 添加按鈕
        self.scroll_To(*self.addBtn)

        # 批註條款_添加按鈕
        self.find_element(*self.PostilGridaddOne).click()

        # 條款代號 NIEO1101 PostilGrid1r0 PostilGrid1r0
        PostilGrid1r0x=self.find_element(*self.PostilGrid1r0)
        ActionChains(self.driver).move_to_element(PostilGrid1r0x).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('NIEO1101')
        # self.take_screenshot_forInput('批註條款相關資訊-原投資代號1')

        # 原投資代號 PostilConsultGrid1r0
        self.find_element(*self.PostilConsultGrid1r0).click()
        self.take_screenshot_forInput('批註條款相關資訊-原投資代號2')

        # 原投資代號名稱 PostilConsultGrid2r0 
        self.find_element(*self.PostilConsultGrid2r0).click()
        self.take_screenshot_forInput('批註條款相關資訊-原投資代號名稱')
        
        # 新投資代號 PostilConsultGrid3r0
        PostilConsultGrid3r0x=self.find_element(*self.PostilConsultGrid3r0)
        ActionChains(self.driver).move_to_element(PostilConsultGrid3r0x).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('E077')
        self.take_screenshot_forInput('批註條款相關資訊-新投資代號')

        # 新投資代號名稱 PostilConsultGrid4r0
        self.find_element(*self.PostilConsultGrid4r0).click()
        self.take_screenshot_forInput('批註條款相關資訊-新投資代號名稱')


        # # 強制執行上一步
        self.driver.execute_script("backBtn = document.querySelector('#riskbutton1');" + "backBtn.onclick();")
	
    def product_NIFA1602(self,product):
        self.switch_default()
        self.switch_frame("fraInterface")
        # 商品錄入按鈕
        self.find_element(*self.Productbtn).click()

        # 險種編碼
        RiskCode2=self.find_element(*self.RiskCode)
        ActionChains(self.driver).move_to_element(RiskCode2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(product["productId"])
        self.take_screenshot_forInput('險種編碼')

        # 進入商品資訊錄入畫面
        self.find_element(*self.back).click()

        # 資產撥回給付方式 
        assectWay2=self.find_element(*self.assectWayNIFU0101)
        ActionChains(self.driver).move_to_element(assectWay2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('0')
        self.take_screenshot_forInput('資產撥回給付方式')

        # 幣別
        self.find_element(*self.CurrencyCode).click()
        self.take_screenshot_forInput('幣別')

        # 年金給付開始年齡 GetYear
        self.find_element(*self.GetYear).send_keys('50')
        self.take_screenshot_forInput('年金給付開始年齡')

        # 年金給付方式
        Payway2=self.find_element(*self.NewPayway)
        ActionChains(self.driver).move_to_element(Payway2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('12')
        self.take_screenshot_forInput('年金給付方式')

        # 第一期保費責任(點擊外圍框框) //html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[4]
        priceCol2=self.find_element(*self.priceCol)
        ActionChains(self.driver).move_to_element(priceCol2).double_click().perform()
        # 第一期保費責任(保費)
        ele=self.find_elements(*self.priceBtn)
        ele2=ele[2]
        ele2.send_keys("1234")
        self.take_screenshot_forInput('保費')

        # 保存
        self.find_element(*self.Savebtn).click()

        # 投資標的受益分配之約定帳戶-帳號 GetBankAccNo
        self.find_element(*self.GetBankAccNo).click()
        self.find_element(*self.GetBankAccNo).send_keys('0123456789')
        self.take_screenshot_forInput('投資標的受益分配之約定帳戶-帳號')

        # 投資標的受益分配之約定帳戶-戶名 GetAccName
        self.find_element(*self.GetAccName).click()
        self.take_screenshot_forInput('投資標的受益分配之約定帳戶-戶名')

        # 投資標的受益分配之約定帳戶-銀行與分行 GetBankCode
        self.find_element(*self.GetBankCode).click()
        self.find_element(*self.GetBankCode).send_keys('8220015')
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('8220015')
        self.find_element(*self.codeselect).click()
        self.take_screenshot_forInput('投資標的受益分配之約定帳戶-銀行與分行')

        # 投資標的代碼 InvestPlanRate1r0
        self.find_element(*self.investCode).click()
        self.find_element(*self.investCode).send_keys("F201")
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value("F201")
        self.find_element(*self.codeselect).click()
        self.take_screenshot_forInput('投資標的代碼')

        # 投資趴數
        self.find_element(*self.InvestPlanRate5r0).click()
        self.find_element(*self.InvestPlanRate5r0).send_keys("100")
        self.take_screenshot_forInput('投資趴數')

        # 每月扣除順序
        self.find_element(*self.InvestPlanRate7r0).click()
        self.find_element(*self.InvestPlanRate7r0).send_keys('1')
        self.take_screenshot_forInput('每月扣除順序')

        # 添加按鈕
        self.find_element(*self.addBtn).click()

        # 條款代號 NIEO1101 PostilGrid1r0 PostilGrid1r0
        self.find_element(*self.PostilGrid1r0).click()
        self.take_screenshot_forInput('批註條款-條款代號')

        # # 強制執行上一步
        self.driver.execute_script("backBtn = document.querySelector('#riskbutton1');" + "backBtn.onclick();")

    def product_NIFU1203(self,product):
        self.switch_default()
        self.switch_frame("fraInterface")
        # 商品錄入按鈕
        self.find_element(*self.Productbtn).click()

        # 險種編碼
        RiskCode2=self.find_element(*self.RiskCode)
        ActionChains(self.driver).move_to_element(RiskCode2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(product["productId"])
        self.take_screenshot_forInput('險種編碼')

        # 進入商品資訊錄入畫面
        self.find_element(*self.back).click()

        # 資產撥回給付方式
        assectWay2=self.find_element(*self.assectWayNIFU0101)
        ActionChains(self.driver).move_to_element(assectWay2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('0')
        self.find_element(*self.assectWayNIFU0101).click()
        self.take_screenshot_forInput('資產撥回給付方式')

        # 保額
        self.find_element(*self.Amnt).click()
        self.find_element(*self.Amnt).send_keys("5000000")
        self.take_screenshot_forInput('保額')

        # 幣別
        self.find_element(*self.CurrencyCode).click()
        self.take_screenshot_forInput('幣別')

        # 第一期保費責任(點擊外圍框框) //html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[4]
        priceCol2=self.find_element(*self.priceCol)
        ActionChains(self.driver).move_to_element(priceCol2).double_click().perform()
        # 第一期保費責任(保費)
        ele=self.find_elements(*self.priceBtn)
        ele2=ele[2]
        ele2.send_keys("1234")
        self.take_screenshot_forInput('保費')

        # 保存
        self.find_element(*self.Savebtn).click()

        # 投資標的受益分配之約定帳戶-帳號 GetBankAccNo
        self.find_element(*self.GetBankAccNo).click()
        self.find_element(*self.GetBankAccNo).send_keys('0123456789')
        self.take_screenshot_forInput('投資標的受益分配之約定帳戶-帳號')

        # 投資標的受益分配之約定帳戶-戶名 GetAccName
        self.find_element(*self.GetAccName).click()
        self.take_screenshot_forInput('投資標的受益分配之約定帳戶-戶名')

        # 投資標的受益分配之約定帳戶-銀行與分行 GetBankCode
        self.find_element(*self.GetBankCode).click()
        self.find_element(*self.GetBankCode).send_keys('8220015')
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('8220015')
        self.find_element(*self.codeselect).click()
        self.take_screenshot_forInput('投資標的受益分配之約定帳戶-銀行與分行')

        # 投資標的代碼 InvestPlanRate1r0
        self.find_element(*self.investCode).click()
        self.find_element(*self.investCode).send_keys("F201")
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value("F201")
        self.find_element(*self.codeselect).click()
        self.take_screenshot_forInput('投資標的代碼')

        # 投資趴數
        self.find_element(*self.InvestPlanRate5r0).click()
        self.find_element(*self.InvestPlanRate5r0).send_keys("100")
        self.take_screenshot_forInput('投資趴數')

        # 每月扣除順序
        self.find_element(*self.InvestPlanRate7r0).click()
        self.find_element(*self.InvestPlanRate7r0).send_keys('1')
        self.take_screenshot_forInput('每月扣除順序')

        # 添加按鈕
        self.find_element(*self.addBtn).click()

        # 條款代號 NIEO1101 PostilGrid1r0 PostilGrid1r0
        self.find_element(*self.PostilGrid1r0).click()
        self.take_screenshot_forInput('批註條款-條款代號')

        # # 強制執行上一步
        self.driver.execute_script("backBtn = document.querySelector('#riskbutton1');" + "backBtn.onclick();")
    
    def product(self,product):
        self.switch_default()
        self.switch_frame("fraInterface")
        # 商品錄入按鈕
        self.find_element(*self.Productbtn).click()

        # 險種編碼
        RiskCode2=self.find_element(*self.RiskCode)
        ActionChains(self.driver).move_to_element(RiskCode2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(product["productId"])
        self.take_screenshot_forInput('險種編碼')

        # 進入商品資訊錄入畫面
        self.find_element(*self.back).click()

        # 判斷 資產撥回給付方式 是否存在
        result=self.isDisplay(*self.assectWayNIFU0101)
        # 資產撥回給付方式
        if(result==True):
            assectWay2=self.find_element(*self.assectWayNIFU0101)
            ActionChains(self.driver).move_to_element(assectWay2).double_click().perform()
            b=self.find_element(*self.codeselect)
            Select(b).select_by_value('0')
            self.find_element(*self.assectWayNIFU0101).click()
            self.take_screenshot_forInput('資產撥回給付方式')
        else:
            print("資產撥回給付方式_沒有在頁面顯示")


        # 判斷 年金給付開始年齡 是否存在
        result=self.isDisplay(*self.GetYear)
        if(result==True):
            # 年金給付開始年齡 GetYear
            self.find_element(*self.GetYear).send_keys('50')
            self.take_screenshot_forInput('年金給付開始年齡')
        else:
            print("年金給付開始年齡_沒有在頁面顯示")

        # 判斷 保額 是否存在
        result=self.isDisplay(*self.Amnt)
        if(result==True):
            # 保額
            self.find_element(*self.Amnt).click()
            self.find_element(*self.Amnt).send_keys("5000000")
            self.take_screenshot_forInput('保額')
        else:
            print("保額_沒有在頁面顯示")

        # 判斷 幣別 是否存在
        result=self.isDisplay(*self.CurrencyCode)
            # 幣別
        if(result==True):
            self.find_element(*self.CurrencyCode).click()
            self.take_screenshot_forInput('幣別')
        else:
            print("幣別_沒有在頁面顯示")

        # 判斷  保費 是否存在
        result=self.isDisplay(*self.priceCol)
        if(result==True):
            # 第一期保費責任(點擊外圍框框) //html/body/form/div[18]/div/div[2]/div/div[2]/div[2]/table/tbody/tr[1]/td[4]
            priceCol2=self.find_element(*self.priceCol)
            ActionChains(self.driver).move_to_element(priceCol2).double_click().perform()
            # 第一期保費責任(保費)
            ele=self.find_elements(*self.priceBtn)
            ele2=ele[2]
            ele2.send_keys("1234")
            self.take_screenshot_forInput('保費')
        else:
            print("保費_沒有在頁面顯示")

        # 保存
        self.find_element(*self.Savebtn).click()

        result=self.isDisplay(*self.GetBankAccNo)
        if(result==True):
            # 投資標的受益分配之約定帳戶-帳號 GetBankAccNo
            self.find_element(*self.GetBankAccNo).click()
            self.find_element(*self.GetBankAccNo).send_keys('0123456789')
            self.take_screenshot_forInput('投資標的受益分配之約定帳戶-帳號')
        else:
            print("投資標的受益分配之約定帳戶-帳號_沒有在頁面顯示")

        result=self.isDisplay(*self.GetAccName)
        if(result==True):
            # 投資標的受益分配之約定帳戶-戶名 GetAccName
            self.find_element(*self.GetAccName).click()
            self.take_screenshot_forInput('投資標的受益分配之約定帳戶-戶名')
        else:
            print("投資標的受益分配之約定帳戶-戶名_沒有在頁面顯示")


        result=self.isDisplay(*self.GetBankCode)
        if(result==True):
            # 投資標的受益分配之約定帳戶-銀行與分行 GetBankCode
            self.find_element(*self.GetBankCode).click()
            self.find_element(*self.GetBankCode).send_keys('8220015')
            b=self.find_element(*self.codeselect)
            Select(b).select_by_value('8220015')
            self.find_element(*self.codeselect).click()
            self.take_screenshot_forInput('投資標的受益分配之約定帳戶-銀行與分行')
        else:
            print("投資標的受益分配之約定帳戶-銀行與分行_沒有在頁面顯示")

        result=self.isDisplay(*self.investCode)
        if(result==True):
            # 投資標的代碼 InvestPlanRate1r0
            self.find_element(*self.investCode).click()
            self.find_element(*self.investCode).send_keys("F201")
            b=self.find_element(*self.codeselect)
            Select(b).select_by_value("F201")
            self.find_element(*self.codeselect).click()
            self.take_screenshot_forInput('投資標的代碼')
        else:
            print("投資標的代碼_沒有在頁面顯示")

        result=self.isDisplay(*self.InvestPlanRate5r0)
        if(result==True):
            # 投資趴數
            self.find_element(*self.InvestPlanRate5r0).click()
            self.find_element(*self.InvestPlanRate5r0).send_keys("100")
            self.take_screenshot_forInput('投資趴數')
        else:
            print("投資趴數_沒有在頁面顯示")

        result=self.isDisplay(*self.InvestPlanRate7r0)
        if(result==True):
            # 每月扣除順序
            self.find_element(*self.InvestPlanRate7r0).click()
            self.find_element(*self.InvestPlanRate7r0).send_keys('1')
            self.take_screenshot_forInput('每月扣除順序')
        else:
            print("每月扣除順序_沒有在頁面顯示")

        # 添加按鈕
        self.find_element(*self.addBtn).click()

        result=self.isDisplay(*self.PostilGrid1r0)
        if(result==True):
            # 條款代號 NIEO1101 PostilGrid1r0 PostilGrid1r0
            self.find_element(*self.PostilGrid1r0).click()
            self.take_screenshot_forInput('批註條款-條款代號')
        else:
            print("批註條款-條款代號_沒有在頁面顯示")

        # # 強制執行上一步
        self.driver.execute_script("backBtn = document.querySelector('#riskbutton1');" + "backBtn.onclick();")


    # --合同其他資訊------------------------------------------------------  

    #  拇指或蓋印章
    ThumbStamp=(By.ID,'ThumbStamp')
    # 分行/分支機搆 BraBankCode2
    BraBankCode2=(By.ID,'BraBankCode2')
    # 簽署人 Signatory
    Signatory=(By.ID,'Signatory')
    # 專案代碼 ProjectCode
    ProjectCode=(By.ID,'ProjectCode')
    # 代轉件 TurnGeneration
    TurnGeneration=(By.ID,'TurnGeneration')
    # 收件日期 ReceiptDate
    ReceiptDate=(By.ID,'ReceiptDate')
    # 受理編號 AcceptNumber
    AcceptNumber=(By.ID,'AcceptNumber')
    # 密戶 SecreatInsured
    SecreatInsured=(By.ID,'SecreatInsured')
    # 增值回饋分享金、健康回饋獎勵金、帳戶價值通知方式 Notification
    Notification=(By.ID,'Notification')
    # 保單製單類型 PolicyPreparationType
    PolicyPreparationType=(By.ID,'PolicyPreparationType')
    # 投保聲明書 InsureDec
    InsureDec=(By.ID,'InsureDec')

    # 保單其他合同資訊
    divOtherInfoImg=(By.ID,'divOtherInfoImg')

    # 合作公司
    PolicyBankCode=(By.ID,'PolicyBankCode')

    def NewInput_otherInfo_input(self):
        self.switch_default()
        self.switch_frame("fraInterface")

        # 移動到保單其他合同資訊
        self.scroll_To(*self.divOtherInfoImg)

        # 合作公司
        self.find_element(*self.PolicyBankCode).click()
        self.take_screenshot_forInput('合作公司')
        

        # 分行/分支機搆 BraBankCode2
        self.find_element(*self.BraBankCode2).click()
        self.take_screenshot_forInput('分行-分支機搆')

        # 簽署人 Signatory
        self.find_element(*self.Signatory).click()
        self.take_screenshot_forInput('簽署人')

        # 專案代碼 ProjectCode
        self.find_element(*self.ProjectCode).click()
        self.take_screenshot_forInput('專案代碼')

        # 代轉件 TurnGeneration
        self.find_element(*self.TurnGeneration).click()
        self.take_screenshot_forInput('代轉件')

        # 收件日期 ReceiptDate
        self.find_element(*self.ReceiptDate).click()
        self.take_screenshot_forInput('收件日期')

        # 受理編號 AcceptNumber
        self.find_element(*self.AcceptNumber).click()
        self.take_screenshot_forInput('受理編號')
            
        # 拇指或蓋印章
        self.find_element(*self.ThumbStamp).click()
        self.take_screenshot_forInput('拇指或蓋印章')

        # 密戶 SecreatInsured
        self.find_element(*self.SecreatInsured).click()
        self.take_screenshot_forInput('密戶')

        # 增值回饋分享金、健康回饋獎勵金、帳戶價值通知方式 Notification
        self.find_element(*self.Notification).click()
        self.take_screenshot_forInput('增值回饋分享金-健康回饋獎勵金-帳戶價值通知方式')

        # 保單製單類型 PolicyPreparationType
        self.find_element(*self.PolicyPreparationType).click()
        self.take_screenshot_forInput('保單製單類型')

        # 投保聲明書 InsureDec
        self.find_element(*self.InsureDec).click()
        self.take_screenshot_forInput('投保聲明書')

   # --繳費資訊------------------------------------------------------

    # 首期繳費方式
    PayMode2=(By.ID,'PayMode2')
    # 續期繳費方式
    SecPayMode2=(By.ID,'SecPayMode2')
    # 圈存交易序號
    ctno=(By.ID,'ctno')
    # 繳款人身份
    StatusOfPayer=(By.ID,'StatusOfPayer')
    # 繳費資訊
    accountImg=(By.ID,'accountImg')
    # 聲明事項
    StatementItemsGrid4r0=(By.ID,'StatementItemsGrid4r0')
    BankLoad2=(By.ID,'BankLoad2')


    def NewInput_paymentInfo_input(self):
        self.switch_default()
        self.switch_frame("fraInterface")

        self.scroll_To(*self.accountImg)

        # 首期繳費方式
        self.find_element(*self.PayMode2).click()
        self.take_screenshot_forInput('首期繳費方式-錯誤')
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('1')
        self.find_element(*self.PayMode2).click()
        self.take_screenshot_forInput('首期繳費方式')
        
        # 續期繳費方式
        self.find_element(*self.SecPayMode2).click()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('1')
        self.take_screenshot_forInput('續期繳費方式')

        # 圈存交易序號
        self.find_element(*self.ctno).click()
        self.take_screenshot_forInput('圈存交易序號')
        
        # 繳款人身份
        self.find_element(*self.StatusOfPayer).click()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('0')
        self.find_element(*self.StatusOfPayer).click()
        self.take_screenshot_forInput('繳款人身份')

        self.scroll_To(*self.BankLoad2)

        # 聲明事項
        StatementItem2=self.find_element(*self.StatementItemsGrid4r0)
        ActionChains(self.driver).move_to_element(StatementItem2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('Y')
        self.take_screenshot_forInput('聲明事項')
    
   # -----受益人錄入----------------------------------------------------------------------------

    # 受益人錄入按鈕
    LBnfButton=(By.ID,'LBnfButton')

    # 所屬被保險人
    BnfSequenceNo=(By.ID,'BnfSequenceNo')
    # BnfSequenceNo=(By.XPATH,'/html/body/form/div[4]/div/table[1]/tbody/tr[1]/td[2]/input[1]')
    # 受益人類型
    BnfType=(By.ID,'BnfType')
    # 同要保人按鈕
    isAppnt=(By.ID,'isAppnt')
    # 同被保險人按鈕
    isInsured=(By.ID,'isInsured')
    # 法定按鈕
    isSameLaw=(By.ID,'isSameLaw')
    # 受益人順位
    BnfGrade=(By.ID,'BnfGrade')
    # 與被保險人關係
    RelationToInsured1=(By.ID,'RelationToInsured1')
    # 原因
    BnfReason=(By.ID,'BnfReason')
    # 姓名
    BnfName=(By.ID,'BnfName')
    # 證件類型
    BnfIDType=(By.ID,'BnfIDType')
    # 證件號碼
    BnfIDNo=(By.ID,'BnfIDNo')
    # 帳號持有人
    BnfAccName=(By.ID,'BnfAccName')
    # 銀行代號
    BnfBankCode=(By.ID,'BnfBankCode')
    # 分行代號
    BnfSubBankCode=(By.ID,'BnfSubBankCode')
    # 銀行帳號
    BnfBankAccNo=(By.ID,'BnfBankAccNo')
    # 出生日期
    BnfBirthday=(By.ID,'BnfBirthday')
    # 國籍
    BnfNativePlace=(By.ID,'BnfNativePlace')
    # 住所
    BnfCityName=(By.ID,'BnfCityName')
    BnfCountyName=(By.ID,'BnfCountyName')
    BnfDetailAddress=(By.ID,'BnfDetailAddress')
    # 外籍住所
    BnfForeignResidence=(By.ID,'BnfForeignResidence')
    # 電話
    BnfPhone=(By.ID,'BnfPhone')
    # 法人註冊設立日期
    IncorporationDate=(By.ID,'IncorporationDate')
    # 保險金指定匯入信託帳號
    InsuranceAccount=(By.ID,'InsuranceAccount')
    # 銀行代碼
    InsuranceBankCode=(By.ID,'InsuranceBankCode')
    # 帳戶持有人
    AccountHolder=(By.ID,'AccountHolder')
    # 帳號
    Account=(By.ID,'Account')
    # 受益比例
    BnfLot=(By.ID,'BnfLot')

    def BenefitPeople(self):
        self.switch_default()
        self.switch_frame("fraInterface")

        # 受益人錄入按鈕
        self.find_element(*self.LBnfButton).click()

        self.switch_targetWindow('受益人資訊錄入')
        self.switch_default()
        self.switch_frame("fraInterface")

        # 所屬被保險人
        self.find_element(*self.BnfSequenceNo).click()
        self.take_screenshot_forInput('所屬被保險人')

        # 受益人類型
        BnfType2=self.find_element(*self.BnfType)
        ActionChains(self.driver).move_to_element(BnfType2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('1')

        self.find_element(*self.BnfSequenceNo).click()
        self.find_element(*self.BnfType).click()
        self.take_screenshot_forInput('受益人類型-身故')

        BnfType2=self.find_element(*self.BnfType)
        ActionChains(self.driver).move_to_element(BnfType2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('2')
        self.take_screenshot_forInput('受益人類型-祝壽')

        # BnfType2=self.find_element(*self.BnfType)
        # ActionChains(self.driver).move_to_element(BnfType2).double_click().perform()
        # b=self.find_element(*self.codeselect)
        # Select(b).select_by_value('5')
        # self.take_screenshot_forInput('受益人類型-年金')
        
        # 受益人類型(身故)
        BnfType2=self.find_element(*self.BnfType)
        ActionChains(self.driver).move_to_element(BnfType2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('1')

        # 同要保人按鈕
        self.find_element(*self.isAppnt).click()
        self.take_screenshot_forInput('同要保人按鈕')

        # 同被保險人按鈕
        self.find_element(*self.isInsured).click()
        self.take_screenshot_forInput('同被保險人按鈕')

        # 法定按鈕
        self.find_element(*self.isSameLaw).click()
        self.take_screenshot_forInput('法定按鈕')
        self.find_element(*self.isAppnt).click()

        # 受益人順位
        BnfGrade2=self.find_element(*self.BnfGrade)
        ActionChains(self.driver).move_to_element(BnfGrade2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('01')
        self.take_screenshot_forInput('受益人順位')

        # 與被保險人關係
        RelationToInsured12=self.find_element(*self.RelationToInsured1)
        ActionChains(self.driver).move_to_element(RelationToInsured12).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('1')
        self.take_screenshot_forInput('與被保險人關係')

        # 原因
        self.find_element(*self.BnfReason).click()
        self.send_keys(self.BnfReason,'test123')
        # self.find_element(*self.BnfReason).send_keys('test')
        self.take_screenshot_forInput('原因')

        # 受益比例
        self.find_element(*self.BnfLot).send_keys('100')

        # 姓名
        self.find_element(*self.BnfName).click()
        self.take_screenshot_forInput('姓名')

        # 證件類型
        self.find_element(*self.BnfIDType).click()
        self.take_screenshot_forInput('證件類型')

        # 證件號碼
        self.find_element(*self.BnfIDNo).click()
        self.take_screenshot_forInput('證件號碼')

        # 帳號持有人
        self.find_element(*self.BnfAccName).send_keys('test')
        self.take_screenshot_forInput('帳號持有人')

        # 銀行代號
        BnfBankCode2=self.find_element(*self.BnfBankCode)
        ActionChains(self.driver).move_to_element(BnfBankCode2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('822')
        self.take_screenshot_forInput('銀行代號')

        # 分行代號
        BnfSubBankCode2=self.find_element(*self.BnfSubBankCode)
        ActionChains(self.driver).move_to_element(BnfSubBankCode2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('8220015')
        self.take_screenshot_forInput('分行代號')

        # 銀行帳號
        self.find_element(*self.BnfBankAccNo).send_keys('0123456789')
        self.take_screenshot_forInput('銀行帳號')

        # 出生日期
        self.find_element(*self.BnfBirthday).click()
        self.take_screenshot_forInput('出生日期')

        # 國籍
        self.find_element(*self.BnfNativePlace).click()
        self.take_screenshot_forInput('國籍')

        self.scroll_To(*self.BnfForeignResidence)

        # 住所
        js='document.getElementById("BnfCityName").removeAttribute("readonly");'
        self.driver.execute_script(js)
        self.find_element(*self.BnfCityName).clear()
        self.find_element(*self.BnfCityName).send_keys('台北市')

        js='document.getElementById("BnfCountyName").removeAttribute("readonly");'
        self.driver.execute_script(js)
        self.find_element(*self.BnfCountyName).clear()
        self.find_element(*self.BnfCountyName).send_keys('信義區')

        js='document.getElementById("BnfDetailAddress").removeAttribute("readonly");'
        self.driver.execute_script(js)
        self.find_element(*self.BnfDetailAddress).clear()
        self.find_element(*self.BnfDetailAddress).send_keys("測試路55號")

        self.take_screenshot_forInput('住所')

        # 外籍住所
        # self.find_element(*self.BnfForeignResidence).send_keys('test road')
        self.find_element(*self.BnfForeignResidence).click()
        self.take_screenshot_forInput('外籍住所')

        # 電話
        # self.find_element(*self.BnfPhone).send_keys('0987654321')
        self.find_element(*self.BnfPhone).click()
        self.take_screenshot_forInput('電話')

        # 法人註冊設立日期
        self.find_element(*self.IncorporationDate).click()
        self.find_element(*self.IncorporationDate).send_keys('2021-11-19')
        self.take_screenshot_forInput('法人註冊設立日期')

        # 保險金指定匯入信託帳號
        self.find_element(*self.InsuranceAccount).click()
        self.take_screenshot_forInput('保險金指定匯入信託帳號')

        # 銀行代碼
        InsuranceBankCode2=self.find_element(*self.InsuranceBankCode)
        ActionChains(self.driver).move_to_element(InsuranceBankCode2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('822')
        self.take_screenshot_forInput('銀行代碼')

        # 帳戶持有人
        self.find_element(*self.AccountHolder).click()
        self.find_element(*self.AccountHolder).send_keys('test')
        self.take_screenshot_forInput('帳戶持有人')

        # 帳號
        self.find_element(*self.Account).click()
        self.find_element(*self.Account).send_keys('0123456789')
        self.take_screenshot_forInput('帳號')

    # 分期定期給付
    InstallPayButton=(By.ID,'InstallPayButton')

    # 身故/完全失能/特定意外傷害第一級失能保險金執行給付方式	
    PayMode=(By.ID,'PayMode')

    # 分期定期給付開始日
    PayStartDate=(By.ID,'PayStartDate')

    # ++ Btn
    InstallmentPayGridaddOne=(By.ID,'InstallmentPayGridaddOne')

    # 保險金種類
    InstallmentPayGrid1r0=(By.ID,'InstallmentPayGrid1r0')

    # 一次性給付
    InstallmentPayGrid3r0=(By.ID,'InstallmentPayGrid3r0')

    # 分期給付-年給付
    InstallmentPayGrid4r0=(By.ID,'InstallmentPayGrid4r0')

    # 給付期間
    InstallmentPayGrid5r0=(By.ID,'InstallmentPayGrid5r0')

    # 簽名查詢
    sqButton=(By.ID,'sqButton')
    # 法定代理人
    riskbutton50=(By.ID,'riskbutton50')

    # 姓名
    LegalName=(By.ID,'LegalName')
    # 與未成年之關係
    RelationTolegal=(By.ID,'RelationTolegal')
    # 證件類型
    LegalIDType=(By.ID,'LegalIDType')
    # 證件號碼
    LegalIDNo=(By.ID,'LegalIDNo')
    # 出生日期
    LegalBrirthDay=(By.ID,'LegalBrirthDay')
    # 國籍
    LegalNativePlace=(By.ID,'LegalNativePlace')
    # 法定代理人類型
    GuardianType=(By.ID,'GuardianType')

    def InstallPay(self):
        # self.switch_default()
        # self.switch_frame("fraInterface")

        # # 分期定期給付按鈕
        # self.find_element(*self.InstallPayButton).click()


        # self.switch_targetWindow('分期定期給付')
        # self.switch_default()
        # self.switch_frame("fraInterface")

        # # 身故/完全失能/特定意外傷害第一級失能保險金執行給付方式
        # PayMode2=self.find_element(*self.PayMode)
        # ActionChains(self.driver).move_to_element(PayMode2).double_click().perform()
        # b=self.find_element(*self.codeselect)
        # Select(b).select_by_value('1')
        # self.take_screenshot_forInput('身故-完全失能-特定意外傷害第一級失能保險金執行給付方式')

        # # 分期定期給付開始日
        # self.find_element(*self.PayStartDate).send_keys('2021-11-19')
        # self.take_screenshot_forInput('分期定期給付開始日')

        # # ++ Btn
        # self.find_element(*self.InstallmentPayGridaddOne).click()

        # # 保險金種類
        # InstallmentPayGrid1r02=self.find_element(*self.InstallmentPayGrid1r0)
        # ActionChains(self.driver).move_to_element(InstallmentPayGrid1r02).double_click().perform()
        # b=self.find_element(*self.codeselect)
        # Select(b).select_by_value('1')
        # self.take_screenshot_forInput('保險金種類')

        # # 一次性給付
        # self.find_element(*self.InstallmentPayGrid3r0).send_keys('100')
        # self.take_screenshot_forInput('一次性給付')

        # # 分期給付-年給付
        # self.find_element(*self.InstallmentPayGrid4r0).send_keys('100')
        # self.take_screenshot_forInput('分期給付-年給付')

        # # 給付期間
        # self.find_element(*self.InstallmentPayGrid5r0).send_keys('1')
        # self.take_screenshot_forInput('給付期間')


        # # # 強制執行上一步
        # self.driver.execute_script("backBtn = document.querySelector('#riskbutton1');" + "backBtn.onclick();")

        # self.switch_window_back()
        # self.switch_targetWindow('扫描件显示')
        # self.switch_default()
        # self.switch_frame("fraInterface")

        # # 簽名查詢
        # self.find_element(*self.sqButton).click()
        # self.switch_targetWindow('影像资料查询')
        # self.take_screenshot_forInput('簽名查詢1')

        # # ----------------------------------------------------------------------
        # self.switch_window_back()
        # self.switch_targetWindow('扫描件显示')
        self.switch_default()
        self.switch_frame("fraInterface")

        # 法定代理人
        self.find_element(*self.riskbutton50).click()
        self.switch_targetWindow('法定代理人')
        # 窗口最大化
        self.driver.maximize_window()

        self.switch_default()
        self.switch_frame("fraInterface")

        # 姓名
        self.find_element(*self.LegalName).click()
        self.find_element(*self.LegalName).send_keys('測試54088')
        self.take_screenshot_forInput('法定代理人-姓名')
        
        # 與未成年之關係
        RelationTolegal2=self.find_element(*self.RelationTolegal)
        ActionChains(self.driver).move_to_element(RelationTolegal2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('1')
        self.find_element(*self.codeselect).click()
        self.take_screenshot_forInput('法定代理人-與未成年之關係')

        # 法定代理人類型
        GuardianType2=self.find_element(*self.GuardianType)
        ActionChains(self.driver).move_to_element(GuardianType2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('0')
        self.take_screenshot_forInput('法定代理人-法定代理人類型')

        # 證件類型     
        self.find_element(*self.LegalIDType).send_keys('0')
        LegalIDType2=self.find_element(*self.LegalIDType)
        ActionChains(self.driver).move_to_element(LegalIDType2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('0')
        self.take_screenshot_forInput('法定代理人-證件類型')

        # 證件號碼
        self.find_element(*self.LegalIDNo).click()
        self.find_element(*self.LegalIDNo).send_keys('C103386751')
        self.take_screenshot_forInput('法定代理人-證件號碼')
        
        # 出生日期 LegalBrirthDay
        self.find_element(*self.LegalBrirthDay).click()
        self.find_element(*self.LegalBrirthDay).send_keys('1964-12-23')
        self.take_screenshot_forInput('法定代理人-出生日期')
        
        # 國籍
        LegalNativePlace2=self.find_element(*self.LegalNativePlace)
        ActionChains(self.driver).move_to_element(LegalNativePlace2).double_click().perform()
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('TW')
        self.take_screenshot_forInput('法定代理人-國籍')


