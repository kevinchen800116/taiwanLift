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

from QueryConfirm import Query_Confirm_forTest
# from U_LoginPage import LoginPage
# from selenium import webdriver

def test_Query(my_fixture):
    Query_Confirm_forTest(my_fixture)

if __name__ == "__main__":
    pytest.main()
