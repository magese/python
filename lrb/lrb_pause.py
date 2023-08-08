import time
import traceback

import openpyxl
from PyQt6.QtCore import pyqtSignal, QObject
from selenium.common import StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

from lrb.common import log, util


class Item:
    row = 0
    id = ''
    log = ''

    def __init__(self, _row, _id):
        self.row = _row
        self.id = _id

    def to_string(self):
        return 'row=%s, id=%s' % (self.row, self.id)


class Excel:
    path = ''
    xlsx = None
    active = None
    max_row = 0
    max_column = 0
    lines: list = None

    def __init__(self, path, xlsx, active, max_row, max_column, lines):
        self.path = path
        self.xlsx = xlsx
        self.active = active
        self.max_row = max_row
        self.max_column = max_column
        self.lines = lines


class LrbPause(QObject):
    __excel_path: str = ''
    __username: str = ''
    __password: str = ''
    excel: Excel = None
    msg = pyqtSignal(str)

    def __init__(self, excel_path, username, password):
        super().__init__()
        self.__excel_path = excel_path
        self.__username = username
        self.__password = password

    # noinspection PyUnresolvedReferences
    def __read_excel(self):
        xlsx = openpyxl.load_workbook(self.__excel_path)
        active = xlsx.active
        max_row = active.max_row
        max_column = active.max_column
        self.msg.emit(log.msg('读取excel最大行数：{}，最大列数：{}', max_row, max_column))

        items = []
        for j in range(1, max_row + 1):
            if str(active.cell(row=j, column=2).value).startswith('success'):
                continue
            item = Item(j, active.cell(row=j, column=1).value)
            items.append(item)
        self.msg.emit(log.msg('共读取待暂停笔记ID {}条', len(items)))
        self.excel = Excel(self.__excel_path, xlsx, active, max_row, max_column, items)

    @staticmethod
    def __pause_note(info, edge):
        retry = 3
        is_paused = True
        while retry > 0:
            try:
                manage = util.wait_for_find_ele(
                    lambda d: d.find_element(by=By.CLASS_NAME, value="manage-list"), edge)
                id_input = util.wait_for_find_ele(
                    lambda d: manage.find_element(by=By.TAG_NAME, value="input"), edge)
                id_input.send_keys('')
                id_input.clear()
                id_input.send_keys(info.id)
                id_input.send_keys(Keys.ENTER)
                time.sleep(1)

                pause_checkbox = util.wait_for_find_ele(
                    lambda d: d.find_element(by=By.CLASS_NAME, value="css-1a4w089"), edge)
                checked = pause_checkbox.get_attribute('data-is-checked')

                pause_switch = util.wait_for_find_ele(
                    lambda d: d.find_element(by=By.CLASS_NAME, value="css-vwjppn"), edge)

                if 'true' == checked:
                    pause_switch.click()
                    is_paused = False
                time.sleep(0.5)
                break
            except StaleElementReferenceException:
                retry -= 1
                time.sleep(0.5)
        return is_paused

    # noinspection PyUnresolvedReferences
    def exec(self):
        self.__read_excel()
        driver = util.prepare(False, self.__username, self.__password)
        size = len(self.excel.lines)
        success = 0
        failure = 0
        result = ''
        start = time.perf_counter()

        for i in range(0, size):
            line = self.excel.lines[i]
            try:
                is_paused = self.__pause_note(line, driver)
                note = '笔记已暂停' if is_paused else '暂停笔记成功'
                self.msg.emit(
                    log.msg('{} => {}：{}', log.loop_msg(i + 1, size, start), note, line.to_string()))

            except BaseException as e:
                traceback.print_exc()
                failure += 1
                result = 'failure:' + repr(e)
                self.msg.emit(log.msg('{} => 暂停笔记异常：{} => {}',
                                      log.loop_msg(i + 1, size, start), result, line.to_string()))

                driver.quit()
                driver = util.prepare(True, self.__username, self.__password)
            else:
                success += 1
                result = 'success-paused' if is_paused else 'success'
                time.sleep(0.5)
            finally:
                start = time.perf_counter()
                self.excel.active.cell(row=line.row, column=2, value=result)
                self.excel.xlsx.save(self.excel.path)

        self.msg.emit(log.msg('截图任务完成，成功{}条，失败{}条', success, failure))
        driver.quit()


# main
# noinspection PyUnresolvedReferences
def main():
    username = 'skiicn_lrb2021@163.com'
    password = 'Mediacom12345'
    filepath = 'C:\\Users\\mages\\Desktop\\暂停创意id.xlsx'
    lp = LrbPause(filepath, username, password)
    lp.msg.connect(lambda m: print(m))
    lp.exec()


if __name__ == '__main__':
    main()
