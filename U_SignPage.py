from time import time
import datetime
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains

# from Zrealtest.測試範例.UAutoConfirmtest import Policy_number

# 現在的位置：承保處理-->個人保單-->簽發保單頁面
class USignPage():

    # 輸入保單號碼
    PublicWorkPoolQueryGrid1r0=(By.ID,'PublicWorkPoolQueryGrid1r0')
    # 點擊查詢按鈕
    publicSearch=(By.ID,'publicSearch')
    # 打勾
    PublicWorkPoolGridChk0=(By.ID,'PublicWorkPoolGridChk0')
    # 簽發保單 name
    signbuttons=(By.CSS_SELECTOR,'input.cssButton')

    # 輸入保單號碼
    def input_SignPage_policyNumbr(self,Policy_number):
        self.find_element(*self.PublicWorkPoolQueryGrid1r0).send_keys(Policy_number)

    # 點擊查詢按鈕
    def click_SignPage_Qbtn(self):
        self.find_element(*self.publicSearch).click()

    # 打勾
    def click_SignPage_checksele(self):
        self.find_element(*self.PublicWorkPoolGridChk0).click()

    def click_SignPage_signbutton(self):
        sign = self.find_elements(*self.signbuttons)
        # 簽發保單按鈕
        sign1 = sign[0]
        ActionChains(self.driver).move_to_element(sign1).click().perform()
        # self.find_element(*self.signbuttons).click()

# 現在的位置：綜合列印-->承保列印-->保單列印頁面
class UPrintSignPage():
    # 輸入框：保單號碼
    ContNo=(By.ID,'ContNo')

    # 查詢按鈕
    PrintSignPage_Qbtn=(By.CSS_SELECTOR,'a.button')
    # PrintSignPage_Qbtn=(By.PARTIAL_LINK_TEXT,'查詢')

    # 列印保單
    # Printbtn=(By.CSS_SELECTOR,'a.button')
    Printbtn=(By.PARTIAL_LINK_TEXT,'列印保單')

    # 水池內選項PoolSelect
    ContGridSel0=(By.ID,'ContGridSel0')

    def input_PrintSignPage_PolicyNumber(self,Policy_number):
        self.find_element(*self.ContNo).send_keys(Policy_number)

    def click_PrintSignPage_Qbtn(self):
        self.find_element(*self.PrintSignPage_Qbtn).click()

    def click_PrintSignPage_PoolFirst(self):
        self.find_element(*self.ContGridSel0).click()

    def click_PrintSignPage_PrintBtn(self):
        self.find_element(*self.Printbtn).click()

# 現在的位置：承保處理-->個人保單-->簽收回條
class USignBackPage():
    ## 回覆狀態 1：同意
    ReplyStatus=(By.ID,'ReplyStatus')
    codeselect=(By.ID,'codeselect')
    ## 輸入保單號碼
    SignRecSlipGrid1r0=(By.ID,'SignRecSlipGrid1r0')
    ## 輸入簽收日期
    SignRecSlipGrid2r0=(By.ID,'SignRecSlipGrid2r0')
    ## 檢核按鈕
    checkButton=(By.ID,'checkButton')
    ## 提交按鈕
    submitButton=(By.ID,'submitButton')

    def select_SignBackPage_agreeBtn(self):
        self.find_element(*self.ReplyStatus).click()
        ## 選擇回覆狀態
        b=self.find_element(*self.codeselect)
        Select(b).select_by_value('1')

    def input_SignBackPage_Policy_number(self,Policy_number):
        self.find_element(*self.SignRecSlipGrid1r0).send_keys(Policy_number)
    
    def input_SignBackPage_SignBackDate(self,personinfo):
        self.find_element(*self.SignRecSlipGrid2r0).send_keys(personinfo["SignBackDate"])

    def click_SignBackPage_checkBtn(self):
        self.find_element(*self.checkButton).click()

    def click_SignBackPage_submitBtn(self):
        self.find_element(*self.submitButton).click()
        # self.switch_window(1)
        # success=self.find_element(*self.contentTD).text
        # return success

# 現在的位置：系統管理-->批次處理任務執行
class TaskPage():
    ## 任務編號
    TaskCode=(By.ID,'TaskCode')
    Taskselect=(By.ID,'codeselect')

    ## 執行日期
    ExeDate=(By.ID,'ExeDate')

    ## 保單號碼
    ContNo=(By.ID,'ContNo')
    ## 業務編號
    BussNo=(By.ID,'BussNo')

    ## 執行
    executeTask=(By.CSS_SELECTOR,'input.cssButton')

    ## 批处理执行完成！
    contentTD=(By.ID,'contentTD')

    def input_TaskPage_TaskCode(self,taskcode):
        self.find_element(*self.TaskCode).send_keys(taskcode)
        b=self.find_element(*self.Taskselect)
        Select(b).select_by_value(taskcode)
        self.find_element(*self.TaskCode).send_keys(Keys.ENTER)

    def input_TaskPage_ExeDate(self,personinfo,taskcode):
        if(taskcode=='NB001'):
            t_str = personinfo["SignBackDate"]

            # 簽回日期先加10天
            d = datetime.datetime.strptime(t_str, '%Y-%m-%d')
            date_ten=(d+datetime.timedelta(days=10)).strftime("%Y-%m-%d")
            
            # 0代表星期一 6代表星期日
            z = datetime.datetime.strptime(date_ten, '%Y-%m-%d').weekday()
            if(z==0 or z==1 or z==2 or z==3 or z==4):
                print(str(z)+":遇平日直接+10")
                print(taskcode+"跑批次的時間:"+date_ten)
                self.find_element(*self.ExeDate).clear()
                self.find_element(*self.ExeDate).send_keys(date_ten)
            elif(z==5):
                print(str(z)+":遇星期六-1")
                ## string => datetime (strptime)
                d = datetime.datetime.strptime(date_ten, '%Y-%m-%d')
                ## datetime => string (strftime)
                date=(d-datetime.timedelta(days=1)).strftime("%Y-%m-%d")
                print(taskcode+"跑批次的時間:"+date)
                self.find_element(*self.ExeDate).clear()
                self.find_element(*self.ExeDate).send_keys(date)
            elif(z==6):
                print(str(z)+":遇星期日-2")
                ## string => datetime (strptime)
                d = datetime.datetime.strptime(date_ten, '%Y-%m-%d')
                ## datetime => string (strftime)
                date=(d-datetime.timedelta(days=2)).strftime("%Y-%m-%d")
                print(taskcode+"跑批次的時間:"+date)
                self.find_element(*self.ExeDate).clear()
                self.find_element(*self.ExeDate).send_keys(date)
        elif(taskcode=='ILP015'):
            t_str = personinfo["SignBackDate"]

            # 簽回日期先加10天
            d = datetime.datetime.strptime(t_str, '%Y-%m-%d')
            date_ten=(d+datetime.timedelta(days=10)).strftime("%Y-%m-%d")

            # 0代表星期一 6代表星期日
            z = datetime.datetime.strptime(date_ten, '%Y-%m-%d').weekday()
            if(z==0 or z==1 or z==2 or z==3 or z==4):
                print(str(z)+":平日")
                ## string => datetime (strptime) 將文字日期轉換為數字
                s = datetime.datetime.strptime(date_ten, '%Y-%m-%d')
                ## datetime => string (strftime)
                date=(s+datetime.timedelta(days=2)).strftime("%Y-%m-%d")
                print(taskcode+"跑批次的時間:"+date)
                self.find_element(*self.ExeDate).clear()
                self.find_element(*self.ExeDate).send_keys(date)
            elif(z==5):
                print(str(z)+":星期六+10-1+4")
                ## string => datetime (strptime) 將文字日期轉換為數字
                s = datetime.datetime.strptime(date_ten, '%Y-%m-%d')
                ## datetime => string (strftime)
                date=(s+datetime.timedelta(days=3)).strftime("%Y-%m-%d")
                print(taskcode+"跑批次的時間:"+date)
                self.find_element(*self.ExeDate).clear()
                self.find_element(*self.ExeDate).send_keys(date)
            elif(z==6):
                print(str(z)+":星期日+10-2+4")
                ## string => datetime (strptime) 將文字日期轉換為數字
                s = datetime.datetime.strptime(date_ten, '%Y-%m-%d')
                ## datetime => string (strftime)
                date=(s+datetime.timedelta(days=2)).strftime("%Y-%m-%d")
                print(taskcode+"跑批次的時間:"+date)
                self.find_element(*self.ExeDate).clear()
                self.find_element(*self.ExeDate).send_keys(date)
# ---------------------------------------------------------------------------------
     # 自行將交費日-25天後 執行CDPA01
    def input_TaskPage_RAW_ExeDate(self,RenewPayDate):
        # print("進入跑批次日期:"+RenewPayDate)
        test_data = datetime.datetime.strptime(RenewPayDate, '%Y-%m-%d')
        date=(test_data-datetime.timedelta(days=25)).strftime("%Y-%m-%d")
        date2=datetime.datetime.strptime(date, '%Y-%m-%d')
        #     ### 將日期帶入  weekday方法  並 判斷是否為假日： 0代表星期一 6代表星期日
        z1=date2
        z=z1.weekday()
        if(z==6):
            print("遇到星期日")
            new_date=(date2+datetime.timedelta(days=1)).strftime("%Y-%m-%d")
        elif(z==5):
            print("遇到星期六")
            new_date=(date2+datetime.timedelta(days=2)).strftime("%Y-%m-%d")
        elif(z != 5 or z != 6):
            print("遇到平常日")
            new_date = date

        print("跑批次日期:"+ new_date)

        self.find_element(*self.ExeDate).clear()
        self.find_element(*self.ExeDate).send_keys(new_date)

    # 自行將自動墊繳日+35天後(2021-12-21) 執行CDPA06
    def input_TaskPage_ExeDate_Plus35(self,APD):
            ## string => datetime (strptime)
            d = datetime.datetime.strptime(APD, '%Y-%m-%d')
            ## datetime => string (strftime)
            date=(d+datetime.timedelta(days=35)).strftime("%Y-%m-%d")
            self.find_element(*self.ExeDate).clear()
            self.find_element(*self.ExeDate).send_keys(date)
# ---------------------------------------------------------------------------------
    # 輸入保單號碼
    def input_TaskPage_PolicyNo(self,Policy_number):
        self.find_element(*self.ContNo).clear()
        self.find_element(*self.ContNo).send_keys(Policy_number)
    # 輸入保單號碼
    def input_TaskPage_BussNo(self,Policy_number):
        self.find_element(*self.BussNo).clear()
        self.find_element(*self.BussNo).send_keys(Policy_number)

    def click_TaskPage_executeTask(self):
        self.find_element(*self.executeTask).click()

    def verify_TaskPage_executeDone(self):
        # 驗證批处理执行完成！
        self.switch_window(1)
        Task_success=self.find_element(*self.contentTD).text
        return Task_success


