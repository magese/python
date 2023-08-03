import time
import traceback

import openpyxl
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

from common import log
from lrb import util


class Item:
    row = 0
    id = ''
    name = ''
    log = ''

    def __init__(self, _row, _id, _name):
        self.row = _row
        self.id = _id
        self.name = _name

    def to_string(self):
        return 'row=%s, id=%s, name=%s' % (self.row, self.id, self.name)


# 读取文件内容
def read_excel(sheet):
    max_row = sheet.max_row
    max_column = sheet.max_column
    log.info('读取excel最大行数：{}，，最大列数：{}', max_row, max_column)

    items = []
    for j in range(1, max_row + 1):
        if sheet.cell(row=j, column=3).value == 'success':
            continue
        item = Item(j, sheet.cell(row=j, column=1).value, sheet.cell(row=j, column=2).value)
        items.append(item)
    return items


# 换名称
def change_note_name(info, edge):
    manage = util.wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="manage-list"), edge)
    id_input = util.wait_for_find_ele(
        lambda d: manage.find_element(by=By.TAG_NAME, value="input"), edge)
    id_input.send_keys('')
    id_input.clear()
    id_input.send_keys(info.id)
    id_input.send_keys(Keys.ENTER)
    time.sleep(0.6)

    edit_div = util.wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="css-ece9u5"), edge)
    edit_a = util.wait_for_find_ele(
        lambda d: edit_div.find_elements(by=By.CLASS_NAME, value="d-link"), edge)
    edit_a[0].click()

    clear_btn = util.wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="css-18cbzsm"), edge)
    clear_btn.click()

    name_input = util.wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="css-1azanbt"), edge)
    name_input.send_keys('')
    name_input.clear()
    name_input.send_keys(info.name)
    time.sleep(0.5)

    finish_btn = util.wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="css-r7neow"), edge)
    finish_btn.click()


# main
driver = util.prepare(False)

filepath = 'C:\\Users\\mages\\Desktop\\创意名称修改.xlsx'
xlsx = openpyxl.load_workbook(filepath)
active = xlsx.active
lines = read_excel(active)
size = len(lines)
log.info('读取文件成功，共读取数据{}条', size)
result = ''

T1 = time.perf_counter()
for i in range(0, size):
    line = lines[i]
    try:
        change_note_name(line, driver)
        log.info('{} => 换笔记名称成功：{}', log.loop_msg(i + 1, size, T1), line.to_string())

    except BaseException as e:
        traceback.print_exc()
        result = 'failure:' + repr(e)
        log.info('{} => 换笔记名称异常：{} => {}', log.loop_msg(i + 1, size, T1), result, line.to_string())

        driver.quit()
        driver = util.prepare(True)

    else:
        result = 'success'
        time.sleep(0.5)
    finally:
        T1 = time.perf_counter()
        active.cell(row=line.row, column=3, value=result)
        xlsx.save(filepath)

log.info('finish')
