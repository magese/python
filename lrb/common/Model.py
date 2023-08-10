import time
import traceback

from PyQt6.QtCore import QThread, pyqtSignal
from selenium.webdriver.edge.webdriver import WebDriver

from lrb.common import log


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


class Lrb(QThread):
    _excel_path: str = ''
    _username: str = ''
    _password: str = ''
    excel: Excel = None
    item = None
    edge: WebDriver = None
    msg = pyqtSignal(str)

    def __init__(self, excel_path, username, password):
        super().__init__()
        self._excel_path = excel_path
        self._username = username
        self._password = password

    # noinspection PyUnresolvedReferences
    def _emit(self, k, *p):
        msg = log.msg(k, *p)
        self.msg.emit(msg)

    def __read_excel(self):
        pass

    def execute(self, action, read_func, page_func, do_func, res_column):
        read_func()
        self.edge = page_func(False, self._username, self._password)

        size = len(self.excel.lines)
        success = 0
        failure = 0
        result = ''
        start = time.perf_counter()

        for i in range(0, size):
            self.item = self.excel.lines[i]
            try:
                flag = do_func()
                note = f'{action}已操作' if flag else f'{action}成功'
                self._emit('{} => {}：{}', log.loop_msg(i + 1, size, start), note, self.item.to_string())

            except BaseException as e:
                traceback.print_exc()
                failure += 1
                result = 'failure:' + repr(e)
                self._emit('{} => {}异常：{} => {}',
                           log.loop_msg(i + 1, size, start), action, result, self.item.to_string())

                self.edge.quit()
                self.edge = page_func
            else:
                success += 1
                result = 'success-completed' if flag else 'success'
                time.sleep(0.5)
            finally:
                start = time.perf_counter()
                self.excel.active.cell(row=self.item.row, column=res_column, value=result)
                self.excel.xlsx.save(self.excel.path)

        self._emit('{}任务完成，成功{}条，失败{}条', action, success, failure)
        self.edge.quit()
