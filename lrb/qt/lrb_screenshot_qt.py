import sys
import time
import traceback
from pathlib import Path

from PyQt6.QtCore import QThread, Qt, pyqtSignal
from PyQt6.QtWidgets import (QWidget, QLineEdit, QGridLayout, QApplication, QFileDialog, QPushButton, QMainWindow,
                             QTextEdit, QMessageBox, QProgressBar)

from common import log
from lrb import util
from lrb.qt.lrb_screenshot import LrbScreenshot


global stop_status


class Worker(QThread):
    logger = None
    excel_path = ''
    save_dir = ''

    status = pyqtSignal(int)
    value_changed = pyqtSignal(float)

    def log(self, k, *p):
        msg = log.msg(k, *p)
        print(msg)
        self.logger.insertPlainText(msg + '\n')

    # noinspection PyUnresolvedReferences
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
            break_flag = False

            if stop_status == 1:
                self.value_changed.emit(0)
                self.status.emit(1)
                self.log('停止执行')
                break
            elif stop_status == 2:
                self.status.emit(2)
                self.log('暂停执行')
                while True:
                    if stop_status == 0:
                        self.status.emit(0)
                        self.log('继续执行')
                        break
                    elif stop_status == 1:
                        break_flag = True
                        break
                    elif stop_status == 2:
                        time.sleep(0.1)
            if break_flag:
                self.log('停止执行')
                self.status.emit(1)
                break

            line = excel.lines[i]
            try:
                save_path = ls.screenshot(drive, line)
                self.log('{} => 保存截图成功：{}', log.loop_msg(i + 1, size, start), save_path)
            except BaseException as e:
                traceback.print_exc()
                failure += 1
                result = 'failure:' + repr(e)
                self.log('{} => 保存截图失败：{}', log.loop_msg(i + 1, size, start), result)
            else:
                success += 1
                result = 'success'
            finally:
                start = time.perf_counter()
                excel.active.cell(row=line.row, column=3, value=result)
                excel.xlsx.save(excel.path)
                self.value_changed.emit((i + 1) / size)

        self.log('截图任务完成，成功{}条，失败{}条', success, failure)
        drive.quit()


class ScreenshotUI(QMainWindow):
    w: QWidget = None
    logger: QTextEdit = None
    _thread: Worker = None
    progress_bar: QProgressBar = None
    excel_btn: QPushButton = None
    excel_edit: QLineEdit = None
    save_dir_btn: QPushButton = None
    save_dir_edit: QLineEdit = None
    stop_btn: QPushButton = None
    pause_btn: QPushButton = None
    start_btn: QPushButton = None

    def __init__(self):
        super().__init__()
        self.init_ui()

    # noinspection PyUnresolvedReferences
    def init_ui(self):
        self.statusBar()
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setRange(0, 1000)

        self._thread = Worker(self)
        self._thread.value_changed.connect(lambda v: self.progress_bar.setValue(int(v * 1000)))
        self._thread.status.connect(self.status)
        self._thread.finished.connect(self.finish)

        w = QWidget()
        self.setCentralWidget(w)

        self.excel_btn = QPushButton('选择excel')
        self.excel_edit = QLineEdit(self)
        self.excel_edit.setReadOnly(True)
        self.excel_btn.clicked.connect(self.excel_select)

        self.save_dir_btn = QPushButton('选择保存位置')
        self.save_dir_edit = QLineEdit(self)
        self.save_dir_edit.setReadOnly(True)
        self.save_dir_btn.clicked.connect(self.save_dir_select)

        self.logger = QTextEdit(self)
        self.logger.setReadOnly(True)
        self.logger.setMinimumHeight(200)
        self.logger.setTextInteractionFlags(Qt.TextInteractionFlag.NoTextInteraction)
        self.logger.viewport().setCursor(Qt.CursorShape.ArrowCursor)
        self._thread.logger = self.logger

        self.stop_btn = QPushButton('停止运行')
        self.pause_btn = QPushButton('暂停运行')
        self.start_btn = QPushButton('开始运行')

        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop)
        self.pause_btn.setEnabled(False)
        self.pause_btn.clicked.connect(self.pause)
        self.start_btn.clicked.connect(self.start)

        grid = QGridLayout()
        grid.setSpacing(10)

        grid.addWidget(self.excel_edit, 1, 0, 1, 4)
        grid.addWidget(self.excel_btn, 1, 4)

        grid.addWidget(self.save_dir_edit, 2, 0, 1, 4)
        grid.addWidget(self.save_dir_btn, 2, 4)

        grid.addWidget(self.stop_btn, 3, 1)
        grid.addWidget(self.pause_btn, 3, 2)
        grid.addWidget(self.start_btn, 3, 3)

        grid.addWidget(self.progress_bar, 4, 0, 1, 5)

        grid.addWidget(self.logger, 5, 0, 8, 5)

        w.setLayout(grid)

        w.setMinimumWidth(500)
        self.setGeometry(300, 300, 650, 100)
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

    def status(self, v):
        if v == 0:
            self.pause_btn.setText('暂停运行')
            self.pause_btn.setCursor(Qt.CursorShape.ArrowCursor)
            self.stop_btn.setEnabled(True)
            self.pause_btn.setEnabled(True)
            self.start_btn.setEnabled(False)
            self.statusBar().showMessage('正在运行……')
        elif v == 1:
            self.excel_btn.setEnabled(True)
            self.save_dir_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)
            self.stop_btn.setText('停止运行')
            self.stop_btn.setCursor(Qt.CursorShape.ArrowCursor)
            self.pause_btn.setEnabled(False)
            self.start_btn.setEnabled(True)
            self.statusBar().showMessage('已停止')
        elif v == 2:
            self.pause_btn.setText('继续运行')
            self.pause_btn.setCursor(Qt.CursorShape.ArrowCursor)
            self.stop_btn.setEnabled(True)
            self.pause_btn.setEnabled(True)
            self.start_btn.setEnabled(False)

    def start(self):
        excel_text = self.excel_edit.text()
        save_text = self.save_dir_edit.text()
        if len(excel_text) == 0 or len(save_text) == 0:
            QMessageBox.warning(self.w, '警告', '请先选择待处理excel文件及保存路径后再开始运行！')
            return

        self._thread.excel_path = self.excel_edit.text()
        self._thread.save_dir = self.save_dir_edit.text()

        global stop_status
        stop_status = 0
        self.excel_btn.setEnabled(False)
        self.save_dir_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.pause_btn.setEnabled(True)
        self.start_btn.setEnabled(False)

        self._thread.start()
        self.statusBar().showMessage('正在运行……')

    def pause(self):
        global stop_status
        if stop_status == 2:
            self.resume()
            return

        stop_status = 2
        self.pause_btn.setEnabled(False)
        self.pause_btn.setText('暂停中...')
        self.pause_btn.setCursor(Qt.CursorShape.WaitCursor)
        self.statusBar().showMessage('暂停中……')

    def resume(self):
        global stop_status
        stop_status = 0
        self.pause_btn.setEnabled(False)
        self.pause_btn.setText('继续中...')
        self.pause_btn.setCursor(Qt.CursorShape.WaitCursor)
        self.statusBar().showMessage('正在运行……')

    def stop(self):
        global stop_status
        stop_status = 1
        self.stop_btn.setEnabled(False)
        self.stop_btn.setText('停止中...')
        self.stop_btn.setCursor(Qt.CursorShape.WaitCursor)
        self.statusBar().showMessage('停止中……')

    def finish(self):
        self.excel_btn.setEnabled(True)
        self.save_dir_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.pause_btn.setEnabled(False)
        self.start_btn.setEnabled(True)
        self.progress_bar.reset()
        self.statusBar().showMessage('已完成')


def main():
    app = QApplication(sys.argv)
    ui = ScreenshotUI()
    ui.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
