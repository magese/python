import sys
import time
import traceback
from pathlib import Path

from PyQt6.QtCore import QThread
from PyQt6.QtWidgets import (QWidget, QLineEdit, QGridLayout, QApplication, QFileDialog, QPushButton, QMainWindow,
                             QTextEdit)

from common import log
from lrb import util
from lrb.qt.lrb_screenshot import LrbScreenshot


class Worker(QThread):

    logger = None
    excel_path = ''
    save_dir = ''

    def log(self, k, *p):
        msg = log.msg(k, *p)
        print(msg)
        self.logger.insertPlainText(msg + '\n')

    def run(self):
        ls = LrbScreenshot(self.excel_path, self.save_dir)
        excel = ls.read_excel()
        self.log('读取excel最大行数：{}，最大列数：{}', excel.max_row, excel.max_column)

        size = len(excel.lines)
        self.log('共读取待截图记录{}条', size)

        if size <= 0:
            return

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
                self.log('{} => save jpg success => {}', log.loop_msg(i + 1, size, start), save_path)
            except BaseException as e:
                traceback.print_exc()
                failure += 1
                result = 'failure:' + repr(e)
                self.log('{} => save jpg failure => {}', log.loop_msg(i + 1, size, start), result)
            else:
                success += 1
                result = 'success'
            finally:
                start = time.perf_counter()
                excel.active.cell(row=line.row, column=3, value=result)
                excel.xlsx.save(excel.path)

        self.log('截图任务完成，成功{}条，失败{}条', success, failure)
        drive.quit()


class ScreenshotUI(QMainWindow):
    w: QWidget = None
    logger: QTextEdit = None
    excel_edit: QLineEdit = None
    save_dir_edit: QLineEdit = None
    _thread: Worker = None

    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self._thread = Worker(self)
        self.statusBar()

        w = QWidget()
        self.setCentralWidget(w)

        excel_btn = QPushButton('选择excel')
        self.excel_edit = QLineEdit(self)
        self.excel_edit.setReadOnly(True)
        excel_btn.clicked.connect(self.excel_select)

        save_dir_btn = QPushButton('选择保存位置')
        self.save_dir_edit = QLineEdit(self)
        self.save_dir_edit.setReadOnly(True)
        save_dir_btn.clicked.connect(self.save_dir_select)

        self.logger = QTextEdit(self)
        self.logger.setReadOnly(True)
        self.logger.setVisible(False)
        self.logger.setMinimumHeight(350)
        self._thread.logger = self.logger

        start_btn = QPushButton('开始运行')
        start_btn.clicked.connect(self.run)

        log_btn = QPushButton('运行日志')
        log_btn.clicked.connect(self.trigger_logger)

        grid = QGridLayout()
        grid.setSpacing(10)

        grid.addWidget(self.excel_edit, 1, 0, 1, 3)
        grid.addWidget(excel_btn, 1, 3)

        grid.addWidget(self.save_dir_edit, 2, 0, 1, 3)
        grid.addWidget(save_dir_btn, 2, 3)

        grid.addWidget(start_btn, 3, 1)
        grid.addWidget(log_btn, 3, 2)

        grid.addWidget(self.logger, 4, 0, 8, 4)

        w.setLayout(grid)

        w.setGeometry(200, 300, 650, 100)
        w.setMinimumWidth(400)
        self.setWindowTitle('小红书截图')

    def excel_select(self):
        w = self.centralWidget()
        home_dir = str(Path.home()) + r'\Desktop'
        fname = QFileDialog.getOpenFileName(w, '选择', home_dir, '文件类型 (*.xlsx)')
        self.excel_edit.setText(fname[0])
        msg = f'选择待处理截图文件：{fname[0]}'
        self.statusBar().showMessage(msg)

    def save_dir_select(self):
        w = self.centralWidget()
        home_dir = str(Path.home()) + r'\Desktop'
        fname = QFileDialog.getExistingDirectory(w, '选择', home_dir)
        self.save_dir_edit.setText(fname)
        msg = f'选择截图保存文件夹：{fname}'
        self.statusBar().showMessage(msg)

    def trigger_logger(self):
        self.logger.setVisible(not self.logger.isVisible())

    def run(self):
        self._thread.excel_path = self.excel_edit.text()
        self._thread.save_dir = self.save_dir_edit.text()
        self._thread.start()


def main():
    app = QApplication(sys.argv)
    ui = ScreenshotUI()
    ui.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
