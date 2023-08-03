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
    log = ''

    def __init__(self, _row, _id):
        self.row = _row
        self.id = _id

    def to_string(self):
        return 'row=%s, id=%s' % (self.row, self.id)


# 读取文件内容
def read_excel(sheet):
    max_row = sheet.max_row
    max_column = sheet.max_column
    log.info('读取excel最大行数：{}，，最大列数：{}', max_row, max_column)

    items = []
    for j in range(1, max_row + 1):
        if sheet.cell(row=j, column=2).value == 'success':
            continue
        item = Item(j, sheet.cell(row=j, column=1).value)
        items.append(item)
    return items


# 暂停笔记
def pause_note(info, edge):
    manage = util.wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="manage-list"), edge)
    id_input = util.wait_for_find_ele(
        lambda d: manage.find_element(by=By.TAG_NAME, value="input"), edge)
    id_input.send_keys('')
    id_input.clear()
    id_input.send_keys(info.id)
    id_input.send_keys(Keys.ENTER)
    time.sleep(0.6)

    pause_checkbox = util.wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="css-1a4w089"), edge)
    checked = pause_checkbox.get_attribute('data-is-checked')

    pause_switch = util.wait_for_find_ele(
        lambda d: d.find_element(by=By.CLASS_NAME, value="css-vwjppn"), edge)

    if 'true' == checked:
        pause_switch.click()
    else:
        log.info('笔记：{}已暂停', info.id)
    time.sleep(0.5)


# main
driver = util.prepare(False)

filepath = 'C:\\Users\\mages\\Desktop\\暂停创意id.xlsx'
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
        pause_note(line, driver)
        log.info('{} => 暂停笔记成功：{}', log.loop_msg(i + 1, size, T1), line.to_string())

    except BaseException as e:
        traceback.print_exc()
        result = 'failure:' + repr(e)
        log.info('{} => 暂停笔记异常：{} => {}', log.loop_msg(i + 1, size, T1), result, line.to_string())

        driver.quit()
        driver = util.prepare(True)
    else:
        result = 'success'
        time.sleep(0.5)
    finally:
        active.cell(row=line.row, column=2, value=result)
        xlsx.save(filepath)

log.info('finish')
