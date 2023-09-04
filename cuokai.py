# This is a sample Python script.
#
# # Press Shift+F10 to execute it or replace it with your code.
# # Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
#
# # coding=utf-8
# # 机房运维

from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.ui import WebDriverWait as Wait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from time import sleep
# import requests
import io
import sys
# import pyttsx3
import winsound
import win32com.client
# from pyttsx3.drivers import sapi5
import datetime
import time
# import paramiko
import os, platform
import socket

"""chrome_options = Options()
chrome_options.debugger_address = "127.0.0.1:4444"

speak_out = win32com.client.Dispatch('SAPI.SPVOICE')

# 隐藏浏览器
# chrome_options.add_argument("--headless")
# driver = webdriver.Ie(executable_path="D:\Python39\Scripts\IEDriverServer.64")
driver = webdriver.Firefox(executable_path="D:\Python39\Scripts\geckodriver", firefox_options=chrome_options)"""
speak_out = win32com.client.Dispatch('SAPI.SPVOICE')
options = webdriver.ChromeOptions()
options.add_argument('disable-infobars')
# 设置下载文件的默认目录
download_directory = 'D:\\新建文件夹'
options.add_experimental_option("prefs", {
    "download.default_directory": download_directory,
    "download.prompt_for_download": False,  # 不显示下载提示框
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})
# #设置 Chrome 浏览器的位置
# chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
# options.binary_location = chrome_path
# 指定 chrome driver 的路径
webdriver_path = r'C:\Windows\System32\chromedriver.exe'
# 创建 Chrome WebDriver 实例，并传入设置的ChromeOptions
driver = webdriver.Chrome(options=options)
driver.maximize_window()
file_name = "d:/text.png"

# 打开10.134.0.186网站
driver.get("http://10.134.0.186")
driver.find_element(By.ID, "username").send_keys("admin")
driver.find_element(By.ID, "upasswd").send_keys("Tzpost2019.")
driver.find_element(By.XPATH, "/html/body/div/div/div/div[1]/div[2]/form/div[3]/div[1]/input").send_keys("admin")
driver.find_element(By.XPATH, "/html/body/div/div/div/div[1]/div[2]/form/div[4]/div/button[1]").click()
handle1 = driver.current_window_handle  # handle1是网络监控标签页
sleep(30)

# 打开集团运维网站并登陆
# driver.execute_script("window.open('http://10.134.100.100:9001')")
driver.execute_script("window.open('http://10.4.228.177')")

handles = driver.window_handles

# 对窗口进行遍历
for handle2 in handles:
    # 筛选新打开的窗口B
    if handle2 != handle1:
        # 切换到新打开的窗口B
        driver.switch_to.window(handle2)  # handle2是运维网

sleep(5)
print(driver.title)
driver.find_element(By.XPATH,
                    "/html/body/form/table/tbody/tr[3]/td[2]/div[2]/div[1]/div[1]/ul/li[1]/span[2]/input").send_keys(
    "tz_ser")
driver.find_element(By.XPATH,
                    "/html/body/form/table/tbody/tr[3]/td[2]/div[2]/div[1]/div[1]/ul/li[2]/span[2]/input").send_keys(
    "js_tz_ser")
driver.find_element(By.ID, "btn").click()
sleep(5)
handle3 = handles[0]


# 播报
def speak(str):
    print(str)
    speak_out.Speak(str)
    winsound.PlaySound(str, winsound.SND_ASYNC)


def yw():
    # 切换窗口
    driver.switch_to.window(handle2)  # handle2 is yunwei shou ye mian

    # 打开待办事项页面
    driver.get("http://10.4.228.177/wf/monitor/ProcesssummaryTodo.do?ispool=1&isindex=1")
    # driver.get("http://10.134.100.100:9001/wf/monitor/ProcesssummaryTodo.do?ispool=1&isindex=1")
    sleep(5)

    # 防止超时自动退出，重新打开集团运维网站
    if driver.title.__contains__("登录入口"):
        driver.get("http://10.4.228.177")
        # driver.get("http://10.134.100.100:9001")
        sleep(5)
        driver.find_element(By.XPATH,
                            "/html/body/form/table/tbody/tr[3]/td[2]/div[2]/div[1]/div[1]/ul/li[1]/span[2]/input").send_keys(
            "tz_ser")
        driver.find_element(By.XPATH,
                            "/html/body/form/table/tbody/tr[3]/td[2]/div[2]/div[1]/div[1]/ul/li[2]/span[2]/input").send_keys(
            "js_tz_ser")
        driver.find_element(By.ID, "btn").click()
        sleep(5)
        driver.get("http://10.4.228.177/wf/monitor/ProcesssummaryTodo.do?ispool=1&isindex=1")
        # driver.get("http://10.134.100.100:9001/wf/monitor/ProcesssummaryTodo.do?ispool=1&isindex=1")
        sleep(5)

    Wait(driver, 60).until(EC.presence_of_element_located((By.ID, "listtable")))
    # logo6=driver.find_element(By.XPATH,"/html/body/form/div/div/div/table/tbody/tr[2]/td[3]/a")
    t_table1 = driver.find_element(By.XPATH, ".//*[@id='listtable']")
    rows1 = t_table1.find_elements(By.TAG_NAME, 'tr')
    if len(rows1) > 1:
        # voice = pyttsx3.init()
        # voice.say('有新的调度令，需要接收')
        # voice.runAndWait()
        speak('有新的调度令，需要尽快接收')
        sleep(5)
        speak('有新的调度令，需要尽快接收')
        sleep(5)
        speak('有新的调度令，需要尽快接收')
        sleep(5)
        speak('有新的调度令，需要尽快接收')
        sleep(5)
        speak('有新的调度令，需要尽快接收')
        filename = time.strftime("%Y%m%d%H%M%S", time.localtime())
        file = "d:/YW/" + filename + ".txt"
        fp = open(file, "w")
        fp.write(filename + '    检测到有新的调度令需要接收' + "\n")
        fp.close()

    sleep(10)


def yw1():
    time2 = datetime.datetime.now()
    # 切换窗口

    driver.switch_to.window(handle2)

    # 打开运维页面
    driver.get("http://10.4.228.177/wf/monitor/Processsummary.do?isindex=1#")
    # driver.get("http://10.134.100.100:9001/wf/monitor/Processsummary.do?isindex=1#")
    sleep(5)

    # 防止超时自动退出，重新打开集团运维网站
    if driver.title.__contains__("登录入口"):
        # driver.get("http://10.134.100.100:9001")
        driver.get("http://10.4.228.177")
        sleep(5)
        driver.find_element(By.XPATH,
                            "/html/body/form/table/tbody/tr[3]/td[2]/div[2]/div[1]/div[1]/ul/li[1]/span[2]/input").send_keys(
            "tz_ser")
        driver.find_element(By.XPATH,
                            "/html/body/form/table/tbody/tr[3]/td[2]/div[2]/div[1]/div[1]/ul/li[2]/span[2]/input").send_keys(
            "js_tz_ser")
        driver.find_element(By.ID, "btn").click()
        sleep(5)

    driver.switch_to.frame(driver.find_element(By.NAME, "_ddajaxtabsiframe-countrydivcontainer"))
    # driver.switch_to_frame(driver.find_element(By.NAME,"_ddajaxtabsiframe-countrydivcontainer"))

    listtable = driver.find_element(By.ID, "listtable")
    Wait(driver, 60).until(EC.presence_of_element_located((By.ID, "listtable")))

    t_table11 = driver.find_element(By.XPATH, ".//*[@id='listtable']")
    trs = t_table11.find_elements(By.TAG_NAME, "tr")
    len1 = len(trs)
    first_tr_i = 2
    end_tr_i = len1 + 1  # 如果新增了行数，或者减少了行数，那么会怎么样

    for i in range(first_tr_i, end_tr_i):

        xPath = "/html/body/form/div/div/div/table/tbody/tr[" + i.__str__() + "]/td[3]/a"

        print("单号", driver.find_element(By.XPATH, xPath).text)

        text111 = driver.find_element(By.XPATH, xPath).text.__str__().find("JS-2023-")

        if text111 == 0 and time2.hour in {8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18}:
            # ru guo shi de
            # jin ru
            driver.find_element(By.XPATH, xPath).click()
            sleep(1)
            # da ying wang zhi
            url = driver.current_url
            # print(url)

            handles = driver.window_handles
            driver.switch_to.window(handles[-1])

            title_d = driver.title
            # print(title_d)
            Wait(driver, 60).until(EC.presence_of_element_located(
                (By.XPATH, "/html/body/form/div[3]/div[3]/table[1]/tbody/tr[10]/td[2]/input")))
            diao_du_name_e = driver.find_element(By.XPATH,
                                                 '/html/body/form/div[3]/div[3]/table[1]/tbody/tr[1]/td[2]/input')
            diao_du_name = diao_du_name_e.get_attribute("value")

            Wait(driver, 60).until(EC.presence_of_element_located(
                (By.XPATH, '/html/body/form/div[3]/div[3]/table[1]/tbody/tr[10]/td[2]/input')))
            endtime_e = driver.find_element(By.XPATH,
                                            '/html/body/form/div[3]/div[3]/table[1]/tbody/tr[10]/td[2]/input')  # 需要添加元素等待

            endtime = endtime_e.get_attribute("value")

            endtime_datetime = datetime.datetime.strptime(endtime, '%Y-%m-%d %H:%M')

            # huo qu localtime

            current_time_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

            current_time = datetime.datetime.strptime(current_time_str, '%Y-%m-%d %H:%M')

            rest_time = endtime_datetime - current_time
            print(rest_time)
            day = rest_time.days
            print(day)
            if day < 2:
                hour = rest_time.seconds // 3600

                minute = (rest_time.seconds - (hour * 3600)) // 60

                str_day = day.__str__()
                str_hour = hour.__str__()
                str_minute = minute.__str__()
                speak("调度令" + diao_du_name + "的办结时间还有" + str_day + "天" + str_hour + "小时" + str_minute + "分钟")
                sleep(5)
                speak("调度令" + diao_du_name + "的办结时间还有" + str_day + "天" + str_hour + "小时" + str_minute + "分钟")
                sleep(5)
                speak("调度令" + diao_du_name + "的办结时间还有" + str_day + "天" + str_hour + "小时" + str_minute + "分钟")

            print("end")

            handles = driver.window_handles
            driver.switch_to.window(handles[1])

            driver.switch_to.frame(driver.find_element(By.NAME, "_ddajaxtabsiframe-countrydivcontainer"))

            sleep(2)
    # driver.get("http://10.4.228.177/wf/monitor/Processsummary.do?isindex=1")


    handles = driver.window_handles
    driver.switch_to.window(handles[1])


def sblist():
    time1 = datetime.datetime.now()
    # 打开10.134.0.186网站的设备列表
    driver.switch_to.window(handle3)
    driver.get("http://10.134.0.186/index.php/Homes/Index.html")
    Wait(driver, 60).until(EC.presence_of_element_located((By.ID, "navbar")))
    driver.find_element(By.NAME, "listOper").click()
    sleep(10)
    t_table = driver.find_element(By.XPATH, ".//*[@id='device-list']/table")
    rows = t_table.find_elements(By.TAG_NAME, 'tr')
    # print(len(rows))
    row1 = 1
    while row1 <= len(rows) - 1:

        t_name = driver.find_element(By.XPATH, ".//*[@id='device-list']/table/tbody/tr[" + str(row1) + "]/td[1]").text
        t_zt = driver.find_element(By.XPATH, ".//*[@id='device-list']/table/tbody/tr[" + str(row1) + "]/td[2]").text
        t_ip = driver.find_element(By.XPATH, ".//*[@id='device-list']/table/tbody/tr[" + str(row1) + "]/td[3]").text
        # t_zl = driver.find_element(By.XPATH,".//*[@id='device-list']/table/tbody/tr["+str(row1)+"]/td[4]").text
        # t_pp = driver.find_element(By.XPATH,".//*[@id='device-list']/table/tbody/tr["+str(row1)+"]/td[5]").text
        print(row1)
        print(t_name + "   " + t_zt + "   " + t_ip)

        if t_zt.__contains__("告警") and time1.hour in {18, 19, 20, 21, 22, 23, 0, 1, 2, 3, 4, 5, 6, 7} and (
                t_name.__contains__("服务器") or t_name.__contains__("汇聚") or
                t_name.__contains__("核心") or t_name.__contains__("云集") or
                t_name.__contains__("小件分拣机") or t_name.__contains__("防火墙") or t_name.__contains__("省市")):

            # 喇叭报警
            # if t_name.__contains__('新锐捷地市核心'):
            #     break
            speak(t_name + '  网点线路中断，请核实')
            sleep(2)
            speak(t_name + '  网点线路中断，请核实')
            sleep(2)
            speak(t_name + '  网点线路中断，请核实')
            sleep(2)
            speak(t_name + '  网点线路中断，请核实')
            sleep(2)
            speak(t_name + '  网点线路中断，请核实')
            sleep(2)
            filename = time.strftime("%Y%m%d%H%M%S", time.localtime())
            file = "d:/YW/" + filename + ".txt"
            fp = open(file, "w")
            fp.write('网点或服务器:' + t_name + t_ip + '  网点线路中断，请核实' + "\n")
            fp.close()
        elif t_zt.__contains__("告警") and time1.hour in {8, 9, 10, 11, 12, 13, 14, 15, 16, 17}:

            speak(t_name + '  告警网点线路中断，请核实')
            sleep(2)
            speak(t_name + '  告警网点线路中断，请核实')
            sleep(2)
            speak(t_name + '  告警网点线路中断，请核实')
            sleep(2)
            speak(t_name + '  告警网点线路中断，请核实')
            sleep(2)
            speak(t_name + '  告警网点线路中断，请核实')
            sleep(2)

            # filename = time.strftime("%Y%m%d%H%M%S", time.localtime())
            # file = "d:/YW/" + filename + ".txt"
            # fp = open(file, "w")
            # fp.write(t_name + ' '+t_ip+'  网点线路中断，请核实' + "\n")
            # fp.close()

        if t_zt.__contains__("正常"):
            row1 = len(rows)
        else:
            #发现故障，判断是否为重要服务器
            if t_ip.__contains__("10.134.0.12") or t_ip.__contains__("10.134.27.110") or t_ip.__contains__(
                    "10.134.27.111") or t_ip.__contains__("10.134.27.112") or t_ip.__contains__(
                "10.134.27.113") or t_ip.__contains__("10.134.27.114") or t_ip.__contains__(
                "10.134.27.115") or t_ip.__contains__("10.134.27.116") or t_ip.__contains__("10.134.27.117"):
                speak(t_name + '  网点线路中断，请核实')
                sleep(2)
                speak(t_name + '  网点线路中断，请核实')
                sleep(2)
                speak(t_name + '  网点线路中断，请核实')
                sleep(2)
                speak(t_name + '  网点线路中断，请核实')
                sleep(2)
                speak(t_name + '  网点线路中断，请核实')
                sleep(2)

            row1 = row1 + 1


class MyClinet:
    def __init__(self):
        # 配置ip和port
        self.host = '10.134.100.243'
        self.port = 20108

    @staticmethod
    def send_message(phone_namber, message):
        # 创建socket对象
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            # 连接服务器

            s.connect(('137.120.3.234', 20108))

            # p = '18936783099'
            # message = "黄河路核心交换机网络节点中断，请核实"
            message = '警告，警告：' + message
            # 格式化
            server_message = f'{phone_namber}:3:{message}\n'.encode("GBK")
            print('server_message')
            print(server_message)
            # 发送消息
            s.sendall(server_message)
            # response_message=s.recv(1024)
            # print('response_message')
            # print(response_message)
            buffer = b''
            response_message = b''
            while True:
                data = s.recv(1024)
                if not data:
                    break
                buffer += data
                if buffer.endswith(b'\n'):
                    response_message = buffer.strip()
                    print(response_message)
                    if response_message == b'tts busy':
                        print("等待接听中---")
                    if response_message == b'ok':
                        print("已接听")

                    if response_message == b'+COLP:TTS speack':
                        print("语音已播送完毕")
                        return True
                    if response_message == b'NO CARRIER\r\nready':
                        print("呼叫失败")
                        return False
                    buffer = b''
            print("超时无响应程序自动退出")
            return False

        except:
            print("连接失败")
            return False
        finally:
            #关闭连接
            s.close()

if __name__ == "__main__":

    while True:
        now = datetime.datetime.now()
        # if  now.hour in{9,12,17,18} and now.minute in {5}:
        if now.minute in {0, 10, 20, 30, 40, 50}:
            print(now)  # 进程活跃提醒
            # speak("整点报时，现在是")
            yw()

            # 每隔60秒检测一次

            sblist()
            sleep(60)

            yw1()
        if now.hour in {9, 10, 11, 12, 13, 14, 15, 16, 17} and now.minute in {0}:
            #speak('现在是'+now.hour.__str__()+'时')
            speak('整点报时')
