# coding=utf-8
import logging

class Log:
    def __init__(self):
        # 日誌名稱
        self.logname="mylog"

    def setMsg(self,level,msg):
        logger=logging.getLogger()

        # 創建一個handeler寫入所有日誌
        fh=logging.FileHandler(r'D:\Users\701489\test\NewZrealtest\Log\mylog.log')
        # 創建一個handeler輸出至控制台
        ch=logging.StreamHandler()
        # 定義日誌輸出格式
        


