import ctypes
import sys
import time
import traceback
from pathlib import Path

import qdarkstyle
from PyQt6.QtCore import QThread, Qt, pyqtSignal, QByteArray
from PyQt6.QtGui import QIcon, QPixmap, QTextCursor, QImage
from PyQt6.QtWidgets import (QWidget, QLineEdit, QGridLayout, QApplication, QFileDialog, QPushButton, QMainWindow,
                             QTextEdit, QMessageBox, QProgressBar)
from qdarkstyle import LightPalette

from lrb.common import log, util
from lrb.screenshot.lrb_screenshot import LrbScreenshot

global stop_status

ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("lrb_screenshot_qt")
ico = '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'


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
        try:
            excel = ls.read_excel()
            self.log('读取excel最大行数：{}，最大列数：{}', excel.max_row, excel.max_column)
        except BaseException as e:
            traceback.print_exc()
            self.log('读取excel失败，请检查文件，错误信息：{}', repr(e))
            return

        size = len(excel.lines)
        self.log('共读取待截图记录{}条', size)

        if size <= 0:
            return

        try:
            drive = util.open_browser()
            drive.maximize_window()
        except BaseException as e:
            traceback.print_exc()
            self.log('打开Edge浏览器失败，请检查驱动。错误信息：{}', repr(e))
            return

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
        self.logger.textChanged.connect(self.scroll_to_end)
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
        self.setGeometry(300, 300, 750, 100)
        data = QByteArray().fromBase64(ico.encode())
        image = QImage()
        image.loadFromData(data)
        icon = QIcon()
        icon.addPixmap(QPixmap(image), QIcon.Mode.Normal, QIcon.State.Off)
        self.setWindowIcon(icon)
        self.setWindowTitle('小红书截图 - V1.0.0  Author: magese@live.cn')

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

    def scroll_to_end(self):
        self.logger.moveCursor(QTextCursor.MoveOperation.End)

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
    """
    python -m nuitka --onefile --mingw64 --standalone --windows-disable-console
    --include-plugin-directory=D:/Develop/python/lrb
    --windows-icon-from-ico=./favicon.ico
    --enable-plugin=pyqt6
    lrb_screenshot_qt.py
    """
    app = QApplication(sys.argv)
    ui = ScreenshotUI()
    ui.setStyleSheet(qdarkstyle.load_stylesheet(qt_api='pyqt6', palette=LightPalette()))
    ui.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
