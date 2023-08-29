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
    stop_status: int = 0
    msg = pyqtSignal(str)
    err = pyqtSignal(str)
    status = pyqtSignal(int)
    value_changed = pyqtSignal(float)

    # noinspection PyUnresolvedReferences
    def _emit(self, k, *p):
        msg = log.msg(k, *p)
        self.msg.emit(msg)

    # noinspection PyUnresolvedReferences
    def _err(self, k, *p):
        msg = log.msg(k, *p)
        self.err.emit(msg)

    def __read_excel(self):
        pass

    # noinspection PyUnresolvedReferences
    def execute(self, action, read_func, page_func, do_func, end_func, res_column, interval_ms=0.5):
        try:
            read_func()
        except BaseException as e:
            traceback.print_exc()
            self._emit('读取Excel异常，请确认文件格式是否正确。错误信息：{}', repr(e))
            self._err('读取Excel异常，请确认文件格式是否正确。错误信息：{}', repr(e))
            return

        try:
            self.edge = page_func(False, self._username, self._password)
        except BaseException as e:
            traceback.print_exc()
            self._emit('打开Edge浏览器失败，请检查驱动。错误信息：{}', repr(e))
            self._err('打开Edge浏览器失败，请检查驱动。错误信息：{}', repr(e))
            return

        size = len(self.excel.lines)
        success = 0
        failure = 0
        result = ''
        start = time.perf_counter()

        for i in range(0, size):
            break_flag = False

            if self.stop_status == 1:
                self.value_changed.emit(0)
                self.status.emit(1)
                self._emit('停止执行')
                break
            elif self.stop_status == 2:
                self.status.emit(2)
                self._emit('暂停执行')
                while True:
                    if self.stop_status == 0:
                        self.status.emit(0)
                        self._emit('继续执行')
                        break
                    elif self.stop_status == 1:
                        break_flag = True
                        break
                    elif self.stop_status == 2:
                        time.sleep(0.1)
            if break_flag:
                self._emit('停止执行')
                self.status.emit(1)
                break

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
                self.edge = page_func(False, self._username, self._password)
            else:
                success += 1
                result = 'success-completed' if flag else 'success'
                time.sleep(interval_ms)
            finally:
                start = time.perf_counter()
                if res_column >= 0:
                    self.excel.active.cell(row=self.item.row, column=res_column, value=result)
                    self.excel.xlsx.save(self.excel.path)
                self.value_changed.emit((i + 1) / size)

        if end_func:
            end_func()
        self._emit('{}任务完成，成功{}条，失败{}条', action, success, failure)
        if self.edge:
            self.edge.quit()
