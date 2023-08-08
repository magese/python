import base64
import time
import traceback

import openpyxl
from selenium.webdriver.common.by import By

from lrb.common import log, util


class Item:
    row = 0
    link = ''
    name = ''

    def __init__(self, _row, _link, _name):
        self.row = _row
        self.link = _link
        self.name = _name

    def to_string(self):
        return 'row=%s, link=%s, name=%s' % (self.row, self.link, self.name)


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


class LrbScreenshot:
    __excel_path = ''
    __save_dir = ''

    def __init__(self, excel_path, save_dir):
        self.__excel_path = excel_path
        self.__save_dir = save_dir

    def read_excel(self):
        xlsx = openpyxl.load_workbook(self.__excel_path)
        active = xlsx.active
        max_row = active.max_row
        max_column = active.max_column

        items = []
        for j in range(1, max_row + 1):
            if str(active.cell(row=j, column=3).value).startswith('success'):
                continue
            item = Item(j, active.cell(row=j, column=1).value, active.cell(row=j, column=2).value)
            items.append(item)
        return Excel(self.__excel_path, xlsx, active, max_row, max_column, items)

    @staticmethod
    def __base64_to_image_file(base64_image, file_path):
        img_data = base64.b64decode(base64_image)
        with open(file_path, 'wb') as f:
            f.write(img_data)

    def screenshot(self, edge, row):
        edge.get(row.link)
        time.sleep(2)
        filename = row.name + '.jpg'
        jpg_path = self.__save_dir + '/' + filename
        note = util.wait_for_find_ele(lambda d: d.find_element(by=By.ID, value='noteContainer'), edge)
        base64_img = note.screenshot_as_base64
        self.__base64_to_image_file(base64_img, jpg_path)
        return jpg_path


# main
def main():
    excel_path = 'C:\\Users\\mages\\Desktop\\LRB笔记截图.xlsx'
    save_dir = 'C:\\Users\\mages\\Desktop\\lrb_screenshot\\'
    ls = LrbScreenshot(excel_path, save_dir)

    excel = ls.read_excel()
    log.info('读取excel最大行数：{}，最大列数：{}', excel.max_row, excel.max_column)

    size = len(excel.lines)
    log.info('共读取待截图记录{}条', size)

    drive = util.open_browser()
    drive.maximize_window()
    success = 0
    failure = 0
    result = ''
    start = time.perf_counter()

    for i in range(0, size):
        line = excel.lines[i]
        try:
            save_path = ls.screenshot(drive, line)
            log.info('{} => 保存截图成功 => {}', log.loop_msg(i + 1, size, start), save_path)
        except BaseException as e:
            traceback.print_exc()
            failure += 1
            result = 'failure:' + repr(e)
            log.info('{} => 保存截图失败 => {}', log.loop_msg(i + 1, size, start), result)
        else:
            success += 1
            result = 'success'
        finally:
            start = time.perf_counter()
            excel.active.cell(row=line.row, column=3, value=result)
            excel.xlsx.save(excel.path)

    log.info('截图任务完成，成功{}条，失败{}条', success, failure)
    drive.quit()
