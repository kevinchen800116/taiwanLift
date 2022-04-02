from warnings import catch_warnings
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import time
import datetime
import traceback
from PIL import ImageGrab


class BasePage(object):
    """
    BasePage封装所有页面都公用的方法,例如driver, url ,FindElement等
    """
    # 初始化driver、url、pagetitle等
    # 实例化BasePage类时，最先执行的就是__init__方法，该方法的入参，其实就是BasePage类的入参。
    # __init__方法不能有返回值，只能返回None
    # self只实例本身，相较于类Page而言。

    def __init__(self, selenium_driver: WebDriver, base_url, pagetitle):
        self.driver = selenium_driver
        self.base_url = base_url
        self.pagetitle = pagetitle

    # 通过title断言进入的页面是否正确。
    # 使用title获取当前窗口title，检查输入的title是否在当前title中，返回比较结果（True 或 False）
    def on_page(self, pagetitle):
        return pagetitle in self.driver.title

    # 打开页面，并校验页面链接是否加载正确
    # 以单下划线_开头的方法，在使用import *时，该方法不会被导入，保证该方法为类私有的。
    def _open(self, url, pagetitle):
        # 使用get打开访问链接地址
        self.driver.get(url)
        self.driver.maximize_window()
        # 使用assert进行校验，打开的窗口title是否与配置的title一致。调用on_page()方法
        assert self.on_page(pagetitle), u"打开开页面失败 %s" % url

    # 定义open方法，调用_open()进行打开链接
    def open(self):
        self._open(self.base_url, self.pagetitle)

    # 重写元素定位方法
    def find_element(self, *loc):
        #        return self.driver.find_element(*loc)
        try:
            # 确保元素是可见的。
            # 注意：以下入参为元组的元素，需要加*。Python存在这种特性，就是将入参放在元组里。
            # WebDriverWait(self.driver,10).until(lambda driver: driver.find_element(*loc).is_displayed())
            # 注意：以下入参本身是元组，不需要加*
            WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(loc))
            return self.driver.find_element(*loc)
        except:
            print
            u"%s 页面中未能找到 %s 元素" % (self, loc)

    def find_elements(self, *loc):
        try:
            WebDriverWait(self.driver, 10).until(
                EC.visibility_of_all_elements_located(loc))
            return self.driver.find_elements(*loc)
        except:
            print
            u"%s 页面中未能找到 %s 元素" % (self, loc)

    def switch_targetWindow(self, title):
        # 獲得當前所有開啟的視窗的控制程式碼
        sreach_windows = self.driver.current_window_handle
        all_handles = self.driver.window_handles

        for handle in all_handles:
            if (handle != sreach_windows):
                self.driver.switch_to.window(handle)
                print("其他窗口:"+self.driver.title)
                if(self.driver.title == title):
                    print(title)
                    break
                else:
                    self.driver.switch_to.window(handle)
                    print("窗口:"+self.driver.title)
            else:
                print('當前頁面title：%s' % self.driver.title)

    # def switch_targetWindow2(self,title):
    #     # 獲得當前所有開啟的視窗的控制程式碼
    #     sreach_windows = self.driver.current_window_handle
    #     all_handles = self.driver.window_handles

    #     for handle in all_handles:
    #         while(handle != sreach_windows):
    #             self.driver.switch_to.window(handle)
    #             print("其他窗口:"+self.driver.title)
    #             if(self.driver.title == title):
    #                 print(title)
    #                 break
    #             else:
    #                 self.driver.switch_to.window(handle)
    #                 print("窗口:"+self.driver.title)
    #                 continue

        # for handle in all_handles:
        #     if (handle != sreach_windows):
        #         self.driver.switch_to.window(handle)
        #         print("其他窗口:"+self.driver.title)
        #         if(self.driver.title == title):
        #             print(title)
        #             break
        #         else:
        #             self.driver.switch_to.window(handle)
        #             print("窗口:"+self.driver.title)
        #     else:
        #         print('當前頁面title：%s'%self.driver.title)

    # #轉換視窗(輸入數字決定要轉到哪個視窗，0是最一開始的視窗)
    def switch_window(self, i):
        windows = self.driver.window_handles
        # self.driver.switch_to.window(windows[1])
        return self.driver.switch_to.window(windows[i])

    def switch_window_back(self):
        windows = self.driver.window_handles
        return self.driver.switch_to.window(windows[0])

    # 處理錯誤訊息彈出的alert，點擊alert的確定
    def switch_to_alert_accept(self):
        WebDriverWait(self.driver, 10).until(EC.alert_is_present())
        return self.driver.switch_to_alert().accept()

    # 重写switch_frame方法
    def switch_frame(self, loc):
        return self.driver.switch_to.frame(loc)

    def switch_default(self):
        return self.driver.switch_to.default_content()

    # 定义script方法，用于执行js脚本，范围执行结果
    def script(self, src):
        self.driver.execute_script(src)

    # 重写定义send_keys方法
    def send_keys(self, loc, vaule, clear_first=True, click_first=True):
        try:
            print(loc)
            print(vaule)
            print("進入輸入")
            # loc = getattr(self,"_%s"% loc)  #getattr相当于实现self.loc(經測試後此方法沒有用)
            if click_first:
                print("進入點擊")
                self.find_element(*loc).click()
            if clear_first:
                print("進入清除")
                self.find_element(*loc).clear()
                self.find_element(*loc).send_keys(vaule)
                print("輸入成功")
        except AttributeError:
            print
            u"%s 页面中未能找到 %s 元素" % (self, loc)

    def take_screenshot(self):
        file_path_test = os.path.abspath(
            os.path.join(os.getcwd(), '../ProcessControl'))
        # print("截圖儲存路徑1;"+file_path_test)
        file_path = os.path.dirname(
            file_path_test) + r"\ProcessControl\Screenshot\Screenshots"
        # print("截圖儲存路徑2:"+file_path)
        # file_path = r'D:\Users\701489\Desktop\Screenshot'
        rq = datetime.datetime.now().strftime("%Y-%m-%d_%H_%M_%S")
        screen_name = file_path + rq + '.png'
        print(screen_name)
        self.driver.get_screenshot_as_file(screen_name)
        print('截圖完成')
        time.sleep(1)

    def take_screenshot_forInput(self, name):
        time.sleep(7)
        clip = ImageGrab.grab()

        file_path_test = os.path.abspath(
            os.path.join(os.getcwd(), '../ProcessControl'))
        # print("截圖儲存路徑1;"+file_path_test)
        file_path = os.path.dirname(
            file_path_test) + r"\ProcessControl\Screenshot\Screenshots"
        # print("截圖儲存路徑2:"+file_path)
        # file_path = r'D:\Users\701489\Desktop\Screenshot'
        rq = datetime.datetime.now().strftime("%Y-%m-%d_%H_%M_%S")
        screen_name = file_path + rq + name + '.png'
        clip.save(screen_name)
        print(name+'_截圖完成')

        # time.sleep(7)
        # file_path_test=os.path.abspath(os.path.join(os.getcwd(),'../ProcessControl'))
        # # print("截圖儲存路徑1;"+file_path_test)
        # file_path=os.path.dirname(file_path_test)+ r"\ProcessControl\Screenshot\Screenshots"
        # # print("截圖儲存路徑2:"+file_path)
        # # file_path = r'D:\Users\701489\Desktop\Screenshot'
        # rq=datetime.datetime.now().strftime("%Y-%m-%d_%H_%M_%S")
        # screen_name = file_path + rq + name + '.png'
        # print(screen_name)
        # self.driver.get_screenshot_as_file(screen_name)
        # print('截圖完成')
        # time.sleep(1)

    def scroll_To(self, *loc):
        element = self.find_element(*loc)
        self.driver.execute_script(
            "arguments[0].scrollIntoView(true);", element)

    # 重定義元素是否可見
    def isDisplay(self, *loc):
        try:
            result = self.driver.find_element(*loc).is_displayed()
            # print("元素有顯示於頁面之中")
            return result
        except Exception:
            # traceback.print_exc()
            print('Unable to locate element')
