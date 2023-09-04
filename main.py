from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import time
import win32com.client

#请确保您的 Chrome WebDriver 文件在系统 PATH 指定的目录下C:\Windows\System32

def login(username, password):
    options = webdriver.ChromeOptions()
    options.add_argument('disable-infobars')

    # 指定 Chrome WebDriver 的路径
    webdriver_path = 'C:\\Windows\\System32\\chromedriver.exe'  # 替换为实际的路径
    # driver = webdriver.Chrome(executable_path=webdriver_path,options=options)
    driver = webdriver.Chrome(options=options)
    driver.get('http://192.168.68.241:8034/user/login?redirect=%2F')

    # 查找用户名输入框并输入用户名
    username_input = driver.find_element(By.XPATH, '//*[@id="username"]')  # 使用适当的选择器
    username_input.send_keys(username)

    # 查找密码输入框并输入密码
    password_input = driver.find_element(By.XPATH, '//*[@id="password"]')  # 使用适当的选择器
    password_input.send_keys(password)

    # 查找登录按钮并点击
    login_button = driver.find_element(By.XPATH, '//*[@id="formLogin"]/div[2]/div/div/span/button')  # 使用适当的选择器
    login_button.click()
    # 停留10秒，以便手动观察页面
    time.sleep(10)
    # 等待一段时间，以便查看登录结果
    driver.implicitly_wait(10)

# 按间距中的绿色按钮以运行脚本。
if __name__ == '__main__':
    while True:
        username = 'admin'  # 替换为实际的用户名
        password = 'Aa3.1415926'  # 替换为实际的密码
        login(username, password)