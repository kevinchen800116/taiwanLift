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

def test_Query(my_fixture):
    Policy_number=Query_Confirm_forTest(my_fixture)
    return Policy_number

ALL_Policy_number=[]
Policy_number=test_Query()
ALL_Policy_number.append(Policy_number)
print(ALL_Policy_number)

# 銀轉(首期/續期)
def test_BankTransfer(my_fixture,Policy_number):
    personinfo=my_fixture[0]
    BankTransfer(Policy_number,personinfo)

### 續期銀轉 (跑批次得到續期繳費通知書) (需手動製磁)
def test_RenewPremiumDate(my_fixture,Policy_number):
    personinfo=my_fixture[0]
    Query_RenewPremiumDate(Policy_number,personinfo)
    

if __name__ == "__main__":
    pytest.main()
