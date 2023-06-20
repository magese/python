from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.edge.options import Options
import time


# 等待查找元素
def wait_for_find_ele(func):
    return WebDriverWait(driver, timeout=10).until(func)


# 打开浏览器
def open_browser():
    options = Options()
    options.add_experimental_option('detach', True)
    options.add_argument(r'--user-data-dir=C:\Users\Magese\AppData\Local\Microsoft\Edge\User Data')
    driver = webdriver.Edge(options=options)
    print('打开Edge成功')
    return driver


# 登录
def login():
    login_tab = wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="css-404bxh"))
    login_tab.click()
    print('点击账号登录按钮成功')

    email_input = wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="css-xno39g"))
    email_input.send_keys('username')
    print('输入邮箱成功')
    time.sleep(2)

    pwd_input = wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="css-cct1ew"))
    pwd_input.send_keys('password')
    print('输入密码成功')

    login_btn = wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="css-wp7z9d"))
    login_btn.click()
    print('登录成功')


# main
driver = open_browser()
driver.get("https://ad.xiaohongshu.com/")
time.sleep(2)
login()
