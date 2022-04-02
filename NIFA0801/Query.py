import os
import sys
# import traceback
import pytest
# import time

# o_path=os.path.abspath(os.path.join(os.getcwd(),'..'))
# sys.path.append(o_path)

### 移動到ProcessControl的資料夾(此動作是為了取得QueryConfirm的py檔案內，的Query_Confirm_forTest()方法)
o_path=os.path.abspath(os.path.join(os.getcwd(),'../ProcessControl'))
sys.path.append(o_path)

from QueryConfirm import BankTransfer, Query_Confirm_forTest, Query_RenewPremiumDate
# from U_LoginPage import LoginPage
# from selenium import webdriver

ALL_Policy_number=[]
ReNew_Policy_number=[]
RenewPayDate=[]

def test_Query(my_fixture):
    Policy_number=Query_Confirm_forTest(my_fixture)
    personinfo=my_fixture[0]
    ALL_Policy_number.append(Policy_number)
    # 儲存續期保單號碼
    if(personinfo["PayMode2"]=="1" and personinfo["SecPayMode2"] == "2"):
        ReNew_Policy_number.append(Policy_number)
        RenewPayDate.append(personinfo["RenewPayDate"])
    print("全部的保單號碼:"+str(ALL_Policy_number))
    print("全部的保單號碼2:"+str(ReNew_Policy_number))
    print("續期扣款日期"+str(RenewPayDate))


# # 銀轉(首期/續期)
# # @pytest.mark.skipif(len(ALL_Policy_number) <= 1,reason="條件為True才會跳過")
# def test_BankTransfer(my_fixture):
#     personinfo=my_fixture[0]

#     # 1代表跳過不執行送件
#     if(personinfo["flag"]=='1'):
#         print("不執行送件")

#     # 2代表將ALL_Policy_number全部送件
#     elif(personinfo["flag"]=='2'):
#         print("全部執行送件")
#         BankTransfer(ALL_Policy_number,personinfo)
#         ALL_Policy_number.clear()

#     # 3代表單筆送件
#     elif(personinfo["flag"]=='3'):
#         print("單筆執行送件")
#         BankTransfer(ALL_Policy_number,personinfo)


# def test_RenewPremiumDate(my_fixture):
#     personinfo=my_fixture[0]
#     if(personinfo["flag"]=='2' and personinfo["PayMode2"]=="1" and personinfo["SecPayMode2"] == "2"):
#         print("多筆送件的續期")
#         Query_RenewPremiumDate(ReNew_Policy_number,personinfo,RenewPayDate)

#     elif(personinfo["flag"]=='3' and personinfo["PayMode2"]=="1" and personinfo["SecPayMode2"] == "2"):
#         print("單筆送件的續期")
#         Query_RenewPremiumDate(ReNew_Policy_number,personinfo,RenewPayDate)
#     else:
#         print("尚未送件")

if __name__ == "__main__":
    pytest.main()
