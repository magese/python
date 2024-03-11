import time
from pathlib import Path
from PIL import Image

from selenium import webdriver
from selenium.common import TimeoutException, StaleElementReferenceException, ElementNotInteractableException
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.wait import WebDriverWait

from lrb.common import log


def compress_image(input_image_path, output_image_path, quality):
    image = Image.open(input_image_path)
    if image.mode == 'RGBA':
        image = image.convert('RGB')
    image.save(output_image_path, optimize=True, quality=quality)


def open_browser():
    options = Options()
    options.add_argument(r'--user-data-dir=' + str(Path.home()) + r'\AppData\Local\Microsoft\Edge\User Data')
    return webdriver.Edge(options=options)


def wait_for_find_ele(func, edge, *timeout):
    retry_time = 3
    if timeout:
        timeout = float(timeout[0])
    else:
        timeout = 10

    while retry_time > 0:
        try:
            return WebDriverWait(edge, timeout=timeout).until(func)
        except StaleElementReferenceException:
            log.info('retry wait for find elements!')
            retry_time -= 1
        time.sleep(0.3)


def search_id(_id, edge):
    retry = 3
    while retry > 0:
        try:
            manage = wait_for_find_ele(
                lambda d: d.find_element(by=By.CLASS_NAME, value="manage-list"), edge)
            id_input = wait_for_find_ele(
                lambda d: manage.find_element(by=By.TAG_NAME, value="input"), edge)
            time.sleep(0.3)
            id_input.send_keys('')
            id_input.clear()
            id_input.send_keys(Keys.CONTROL, 'a')
            time.sleep(0.3)
            id_input.send_keys(_id)
            id_input.send_keys(Keys.ENTER)
            break
        except (ElementNotInteractableException, StaleElementReferenceException):
            retry -= 1
            time.sleep(0.5)


def click_edit(edge):
    retry = 3
    while retry > 0:
        try:
            edit_div = wait_for_find_ele(
                lambda d: d.find_element(by=By.CLASS_NAME, value="css-ece9u5"), edge)
            edit_a = wait_for_find_ele(
                lambda d: edit_div.find_elements(by=By.TAG_NAME, value="a"), edge)
            edit_a[0].click()
            break
        except (ElementNotInteractableException, StaleElementReferenceException):
            retry -= 1
            time.sleep(0.5)


def login(username, password, edge):
    try:
        WebDriverWait(edge, timeout=3).until(lambda d: d.find_element(by=By.CLASS_NAME, value="action-icon-group"))
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


# 切换到单元id页面
def switch_unit_id_page(edge):
    menu = wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="d-new-menu__inner"), edge)
    menu_items = wait_for_find_ele(
        lambda d: menu.find_elements(by=By.CLASS_NAME, value="d-menu-item"), edge)
    menu_items[1].click()

    tabs_header = wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="d-tabs-headers-wrapper"), edge)
    tabs = wait_for_find_ele(
        lambda d: tabs_header.find_elements(by=By.CLASS_NAME, value="d-tabs-header"), edge)
    tabs[1].click()

    select = wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="d-select-wrapper"), edge)
    select.click()

    options_wrapper = wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="d-options-wrapper"), edge)
    options = wait_for_find_ele(
        lambda d: options_wrapper.find_elements(by=By.CLASS_NAME, value="d-option-name"), edge)
    options[1].click()


# 到主页
def home_page(reopen, username, password):
    edge = open_browser()
    log.info('重新打开Edge成功' if reopen else '打开Edge成功')

    edge.get("https://ad.xiaohongshu.com/")
    time.sleep(1)

    login(username, password, edge)
    log.info('重新登录账号成功' if reopen else '登录账号成功')
    time.sleep(1)
    return edge


# 到创意页面
def creative_page(reopen, username, password):
    edge = home_page(reopen, username, password)

    switch_note_id_page(edge)
    log.info('重新切换到创意页面成功' if reopen else '切换到创意页面成功')
    time.sleep(1)
    return edge


# 到单元页面
def unit_page(reopen, username, password):
    edge = home_page(reopen, username, password)

    switch_unit_id_page(edge)
    log.info('重新切换到单元页面成功' if reopen else '切换到单元页面成功')
    time.sleep(1)
    return edge
