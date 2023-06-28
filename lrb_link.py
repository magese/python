from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.keys import Keys
import time
import openpyxl


class Item:
    row = 0
    id = ''
    link = ''
    log = ''

    def __init__(self, _row, _id, _link):
        self.row = _row
        self.id = _id
        self.link = _link

    def to_string(self):
        return 'row=%s, id=%s, link=%s' % (self.row, self.id, self.link)


# 等待查找元素
def wait_for_find_ele(func):
    ele = WebDriverWait(driver, timeout=10).until(func)
    time.sleep(0.5)
    return ele


# 打开浏览器
def open_browser():
    options = Options()
    options.add_experimental_option('detach', True)
    options.add_argument(r'--user-data-dir=C:\Users\Magese\AppData\Local\Microsoft\Edge\User Data')
    return webdriver.Edge(options=options)


# 登录
def login(username, password):
    login_tab = wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="css-404bxh"))
    login_tab.click()

    email_input = wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="css-xno39g"))
    email_input.send_keys(username)

    pwd_input = wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="css-cct1ew"))
    pwd_input.send_keys(password)

    login_btn = wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="css-wp7z9d"))
    login_btn.click()


# 切换页面
def switch_page():
    menu = wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="d-new-menu__inner"))
    menu_items = wait_for_find_ele(
        lambda d: menu.find_elements(by=By.CLASS_NAME, value="d-menu-item"))
    menu_items[1].click()

    tabs_header = wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="d-tabs-headers-wrapper"))
    tabs = wait_for_find_ele(
        lambda d: tabs_header.find_elements(by=By.CLASS_NAME, value="d-tabs-header"))
    tabs[2].click()

    select = wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="d-select-wrapper"))
    select.click()

    options_wrapper = wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="d-options-wrapper"))
    options = wait_for_find_ele(
        lambda d: options_wrapper.find_elements(by=By.CLASS_NAME, value="d-option-name"))
    options[1].click()


# 读取文件内容
def read_excel(sheet):
    max_row = sheet.max_row
    max_column = sheet.max_column
    print('最大行数：', max_row, '，最大列数：', max_column)

    items = []
    for j in range(1, max_row + 1):
        if sheet.cell(row=j, column=3).value == 'success':
            continue
        item = Item(j, sheet.cell(row=j, column=1).value, sheet.cell(row=j, column=2).value)
        items.append(item)
    return items


# 换链接
def change_link(info):
    manage = wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="manage-list"))
    id_input = wait_for_find_ele(
        lambda d: manage.find_element(by=By.TAG_NAME, value="input"))
    id_input.clear()
    id_input.send_keys(info.id)
    id_input.send_keys(Keys.ENTER)


# main
driver = open_browser()
print('打开Edge成功')
driver.get("https://ad.xiaohongshu.com/")
time.sleep(2)

login('un', 'pw')
print('登录账号成功')
time.sleep(2)

switch_page()
print('切换页面成功')
time.sleep(2)

filepath = 'C:\\Users\\Magese\\Desktop\\lrb_demo.xlsx'
xlsx = openpyxl.load_workbook(filepath)
active = xlsx.active
lines = read_excel(active)
size = len(lines)
print('读取文件成功，共读取数据', size, "条")

for i in range(0, size):
    line = lines[i]
    try:
        change_link(line)
        print(i + 1, '/', size, '-', format((i + 1) / size * 100, '.2f'), '% => ', line.to_string())
    except RuntimeError:
        print('换链接异常：', line.to_string())
        active.cell(row=line.row, column=3, value='failure')
    else:
        active.cell(row=line.row, column=3, value='success')

xlsx.save(filepath)
print('finish')
