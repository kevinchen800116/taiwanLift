from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import time

# 現在的位置：保費-->續期作業-->通知書查詢
class Renew():
    # 輸入保單號碼
    ContNo=(By.ID,'ContNo')
    # 點擊查詢
    Qbtn=(By.XPATH,'/html/body/form/a')
    # 得到續期繳費通知書的催告日
    NoticePremGrid6r0=(By.ID,'NoticePremGrid6r0')
    # 得到催告通知書的自動墊繳起日
    NoticePremGrid7r1=(By.ID,'NoticePremGrid7r1')


    def input_Renew_PolicyNumber(self,Policy_number):
        # self.find_element(*self.ContNo).send_keys(Policy_number["Policy_number0"])
        self.find_element(*self.ContNo).clear()
        self.find_element(*self.ContNo).send_keys(Policy_number)

    def click_Renew_Qbtn(self):
        self.find_element(*self.Qbtn).click()

    # 得到續期繳費通知書的催告日
    def get_Renew_NoticeDate(self):
        NoticeDate=self.find_element(*self.NoticePremGrid6r0).get_attribute('value')
        return NoticeDate
    # 得到催告通知書的自動墊繳起日
    def get_Renew_APD(self):
        APD=self.find_element(*self.NoticePremGrid7r1).get_attribute('value')
        return APD

# 現在的位置：保費-->銀行轉帳-->授權書建檔
class Bank_Authorization():
    # 新增授權書(switch window)
    NewAuthorBook=(By.PARTIAL_LINK_TEXT,'新增授權書')
    # 授權書條碼
    AuthorBar=(By.ID,'AuthorBar')
    # 授權書類別
    AuthorType=(By.ID,'AuthorType')
    # 選擇框
    codeselect=(By.ID,'codeselect')
    # 申請日期
    AppDate=(By.ID,'AppDate')
    # 業務員登錄證字號
    AgentCode=(By.ID,'AgentCode')

    # 授權人與保單關係
    Renewaln2Grid5r0=(By.ID,'Renewaln2Grid5r0')

    # 是否檢附關係説明文件
    Renewaln2Grid6r0=(By.ID,'Renewaln2Grid6r0')

    # # 簽名查閱
    SignCheck=(By.XPATH,'/html/body/form/a[1]')
    # # 檢核
    CheckBTN1=(By.XPATH,'/html/body/form/a[2]')
    # # 保存
    SaveBTN1=(By.XPATH,'/html/body/form/a[3]')

    # 授權書幣別
    AuthorCurrency=(By.ID,'AuthorCurrency')
    # 銀行代碼
    BankCode=(By.ID,'BankCode')

    # # 檢核2
    CheckBTN2=(By.XPATH,'/html/body/form/a[4]')
    # # 應付未付
    PayBTN=(By.XPATH,'/html/body/form/a[5]')
    # # 保存2
    SaveBTN2=(By.XPATH,'/html/body/form/a[6]')
    # # 確認
    # /html/body/form/a[8]

    # 點擊新增授權書按鈕
    def click_BankAZ_NewAuthorBook_Btn(self):
        self.find_element(*self.NewAuthorBook).click()
        self.switch_window(1)
        t1=self.driver.title
        print("目前窗口:"+t1)

    # 輸入授權書條碼
    def input_BankAZ_AuthorBar(self,Policy_number):
        self.find_element(*self.AuthorBar).send_keys("ML"+Policy_number)

    # 點擊授權書訊息內的 簽名查閱、檢核、保存
    def click_BankAZ_SaveBTN1(self):
        self.find_element(*self.SignCheck).click()
        self.find_element(*self.CheckBTN1).click()

        # 切換到彈窗
        alert = self.driver.switch_to.alert
        # 列印彈窗顯示的資訊:Alert Message
        print(alert.text) 
        # 接受
        alert.accept()

        self.find_element(*self.SaveBTN1).click()

    # 點擊授權書詳細訊息內的 檢核、應付未付、保存
    def click_BankAZ_SaveBTN2(self):
        self.find_element(*self.CheckBTN2).click()

         # 切換到彈窗
        alert = self.driver.switch_to.alert
        # 列印彈窗顯示的資訊:Alert Message
        print(alert.text) 
        # 接受
        alert.accept()

        self.find_element(*self.PayBTN).click()
        time.sleep(10)
        self.find_element(*self.SaveBTN2).click()
        self.switch_window_back()
# 現在的位置：保費-->銀行轉帳-->核印送件作業
class Bank_Delivery():
    # 掃描條碼欄
    AuthorBar1=(By.ID,'AuthorBar1')

    # 查詢結果清單 (打勾)
    checkAllNDResultGrid2=(By.XPATH,'/html/body/form/div[2]/table/tbody/tr/td/span/div/div[3]/div/table/thead/tr/th[1]/div/input')

    # 送件按鈕
    # subBtn=(By.XPATH,'/html/body/form/a[2]')
    DeliveryBtn=(By.CSS_SELECTOR,'a.button')

    PDF=(By.CSS_SELECTOR,'a.ICOhover')

    okBtn=(By.XPATH,'/html/body/div[6]/div[3]/a[1]')

    def input_BD_AuthorBar1(self,Policy_number):
        for i in range(len(Policy_number)):
            self.find_element(*self.AuthorBar1).send_keys("ML"+Policy_number[i])
        time.sleep(1)

    def click_BD_CheckBtn(self):
        self.find_element(*self.checkAllNDResultGrid2).click()
    
    # 下載PDF功能 
    def click_BD_DeliveryBtn(self): 
        DelBtn=self.find_elements(*self.DeliveryBtn)
        DelBtn[1].click()
        # 切換到彈窗
        alert = self.driver.switch_to.alert
        # 列印彈窗顯示的資訊:Alert Message
        print(alert.text) 
        # 接受
        alert.accept()

        time.sleep(2)

        self.switch_window(2)
        print("我是列印頁面:"+self.driver.title)


        PdfDownload=self.find_elements(*self.PDF)
        PdfDownload[1].click()

        self.find_element(*self.okBtn).click()

        time.sleep(2)
        # self.switch_window_back()

# 現在的位置：保費-->銀行轉帳-->核印送核作業
class Bank_Approval():
    # 掃描條碼欄
    AuthorBar1=(By.ID,'AuthorBar1')

    # 查詢結果清單 (打勾)
    # checkAllCPSendGrid=(By.ID,'checkAllCPSendGrid')
    checkAllCPSendGrid=(By.XPATH,'/html/body/form/div[2]/table/tbody/tr/td/span/div/div[3]/div/table/thead/tr/th[1]/div/input')
    
    # 送核
    ApproveBtn=(By.CSS_SELECTOR,'a.button')

    PDF=(By.CSS_SELECTOR,'a.ICOhover')

    okBtn=(By.XPATH,'/html/body/div[6]/div[3]/a[1]/span/span[1]')

    def input_BA_AuthorBar1(self,Policy_number):
        for i in range(len(Policy_number)):
            self.find_element(*self.AuthorBar1).send_keys("ML"+Policy_number[i])

    def click_BA_checkBtn(self):
        self.find_element(*self.checkAllCPSendGrid).click()

    def click_BA_ApproveBtn(self):
        AppBtn=self.find_elements(*self.ApproveBtn)
        AppBtn[1].click()
        # 切換到彈窗
        alert = self.driver.switch_to.alert
        # 列印彈窗顯示的資訊:Alert Message
        print(alert.text) 
        # 接受
        alert.accept()

        time.sleep(2)
        # 轉換網頁內容
        self.switch_window(1)
        # 確定跳轉視窗的名稱
        print("我是列印頁面2:"+self.driver.title)
        # self.switch_window(-1)
        # print("我是-1"+self.driver.title)

        PdfDownload=self.find_elements(*self.PDF)
        PdfDownload[1].click()

        self.find_element(*self.okBtn).click()

        time.sleep(2)
        self.switch_window_back()
# 現在的位置：保費-->銀行轉帳-->核印回件作業
class Bank_Return():
    # 掃描條碼欄
    AuthorBar=(By.ID,'AuthorBar')

    # 銀行核印結果
    BankSealResult=(By.ID,'BankSealResult')
    codeselect=(By.ID,'codeselect')
    # 01

    # 全部提交
    # /html/body/form/a[1]
    submitBtn=(By.CSS_SELECTOR,'a.button')


    # BankType回件類型
    BankType=(By.ID,'BankType')


    # 讀取檔案按鈕
    # /html/body/form/a[2]
    ReadBtn=(By.CSS_SELECTOR,'a.button')
    # 授權書幣別
    AuthorCurrency=(By.ID,'AuthorCurrency')

    file=(By.ID,'file')

    def input_BR_AuthorBar(self,Policy_number):
        self.find_element(*self.AuthorBar).send_keys("ML"+Policy_number)

    def select_BR_result(self):
        # 輸入銀行核印結果
        self.find_element(*self.BankSealResult).send_keys('01')
        time.sleep(1)
        # 選擇核印成功
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('01')
        self.find_element(*self.BankSealResult).send_keys(Keys.ENTER)

    def click_BF_submitBtn(self):
        subBtn=self.find_elements(*self.submitBtn)
        subBtn[0].click()

    def click_BF_saveBtn(self,personinfo):
        # 選擇回件類型
        self.find_element(*self.BankType).send_keys("02")
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('02')
        self.find_element(*self.BankType).send_keys(Keys.ENTER)

        # 選擇幣別AuthorCurrency
        self.find_element(*self.AuthorCurrency).click()
        self.find_element(*self.AuthorCurrency).send_keys(personinfo["AuthorCurrency"])
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(personinfo["AuthorCurrency"])
        self.find_element(*self.AuthorCurrency).send_keys(Keys.ENTER)
        self.find_element(*self.AuthorBar).click()

        # 選擇銀行代碼personinfo["bankcode"]
        # javascript強制執行input標籤的ondblclick()方法
        self.driver.execute_script("BankCode = document.getElementById('BankCode');" + "BankCode.ondblclick();")

        # self.find_element(*self.BankCode).send_keys(personinfo["bankcode"])  # 這個會找不到 報錯:'NoneType' object has no attribute 'send_keys'
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value(personinfo["bankcode"])

        # 讀取檔案按鈕
        ReadBtns=self.find_elements(*self.ReadBtn)
        ReadBtns[1].click()

        # 保存 強制執行保存按鈕的onclick()方法
        self.driver.execute_script("saveBtn = document.getElementsByClassName('button');" + "saveBtn[2].onclick();")

        time.sleep(10)

    def upload_BF_file(self,personinfo):
        # 修改檔案
        # 陽信銀行
        if(personinfo["bankcode"]=="1080014"):
            f1 = open(r'D:\Users\701489\Desktop\PDF\94A.txt', 'r')
            content1 = f1.read()
            content3=content1.replace("A M","A0M")
            file = open(r'D:\Users\701489\Desktop\PDF\94AR.txt', 'w')
            file.write(content3)
            f1.close
            print("修改成功")

            # 4.定位上傳檔案按鈕
            upfile = self.find_element(*self.file)
            # 5.使用send_keys方法上傳檔案
            upfile.send_keys(r'D:\Users\701489\Desktop\PDF\94AR.txt')

        # 玉山銀行
        elif(personinfo["bankcode"]=="8080015"):
            f1 = open(r'D:\Users\701489\Desktop\PDF\808bank.txt', 'r')
            content1 = f1.read()
            content3=content1.replace("N ","N0")
            file = open(r'D:\Users\701489\Desktop\PDF\808bankr.txt', 'w')
            file.write(content3)
            f1.close
            file.close

            f1 = open(r'D:\Users\701489\Desktop\PDF\808bankr.txt', 'r')
            content1 = f1.read()
            content3=content1.replace("N ","N0")
            file = open(r'D:\Users\701489\Desktop\PDF\808bankr.txt', 'w')
            file.write(content3)
            f1.close
            file.close

            f1 = open(r'D:\Users\701489\Desktop\PDF\808bankr.txt', 'r')
            content1 = f1.read()
            content3=content1.replace("N ","N0")
            file = open(r'D:\Users\701489\Desktop\PDF\808bankr.txt', 'w')
            file.write(content3)
            f1.close
            file.close
            print("修改成功")

            # 4.定位上傳檔案按鈕
            upfile = self.find_element(*self.file)
            # 5.使用send_keys方法上傳檔案
            upfile.send_keys(r'D:\Users\701489\Desktop\PDF\808bankr.txt')

    # 電子檔尚未做

