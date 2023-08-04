import time

from selenium import webdriver
from selenium.common import TimeoutException, StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.wait import WebDriverWait

from common import log


# 打开浏览器
def open_browser():
    options = Options()
    options.add_argument(r'--user-data-dir=C:\Users\Magese\AppData\Local\Microsoft\Edge\User Data')
    return webdriver.Edge(options=options)


# 等待查找元素
def wait_for_find_ele(func, edge):
    retry_time = 3

    while retry_time > 0:
        try:
            return WebDriverWait(edge, timeout=10).until(func)
        except StaleElementReferenceException:
            log.info('retry wait for find elements!')
            retry_time -= 1
        time.sleep(0.3)


# 登录
def login(username, password, edge):
    try:
        WebDriverWait(edge, timeout=2).until(lambda d: d.find_element(by=By.CLASS_NAME, value="action-icon-group"))
    except TimeoutException:
        login_tab = wait_for_find_ele(
            lambda d: d.find_element(by=By.CLASS_NAME, value="css-404bxh"), edge)
        login_tab.click()

        email_input = wait_for_find_ele(
            lambda d: d.find_element(by=By.CLASS_NAME, value="css-xno39g"), edge)
        email_input.send_keys(username)

        pwd_input = wait_for_find_ele(
            lambda d: d.find_element(by=By.CLASS_NAME, value="css-cct1ew"), edge)
        pwd_input.send_keys(password)

        login_btn = wait_for_find_ele(
            lambda d: d.find_element(by=By.CLASS_NAME, value="css-wp7z9d"), edge)
        login_btn.click()
    else:
        log.info('当前已登录')


# 切换到笔记id页面
def switch_note_id_page(edge):
    menu = wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="d-new-menu__inner"), edge)
    menu_items = wait_for_find_ele(
        lambda d: menu.find_elements(by=By.CLASS_NAME, value="d-menu-item"), edge)
    menu_items[1].click()

    tabs_header = wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="d-tabs-headers-wrapper"), edge)
    tabs = wait_for_find_ele(
        lambda d: tabs_header.find_elements(by=By.CLASS_NAME, value="d-tabs-header"), edge)
    tabs[2].click()

    select = wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="d-select-wrapper"), edge)
    select.click()

    options_wrapper = wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="d-options-wrapper"), edge)
    options = wait_for_find_ele(
        lambda d: options_wrapper.find_elements(by=By.CLASS_NAME, value="d-option-name"), edge)
    options[1].click()


# 准备
def prepare(reopen):
    edge = open_browser()
    log.info('重新打开Edge成功' if reopen else '打开Edge成功')

    edge.get("https://ad.xiaohongshu.com/")
    time.sleep(1)

    login('', '', edge)
    log.info('重新登录账号成功' if reopen else '登录账号成功')
    time.sleep(1)

    switch_note_id_page(edge)
    log.info('重新切换页面成功' if reopen else '切换页面成功')
    time.sleep(1)
    return edge
