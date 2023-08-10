import ctypes
import sys
from pathlib import Path

import qdarkstyle
from PyQt6.QtCore import Qt, QByteArray
from PyQt6.QtGui import QIcon, QPixmap, QTextCursor, QImage
from PyQt6.QtWidgets import (QWidget, QLineEdit, QGridLayout, QApplication, QFileDialog, QPushButton, QMainWindow,
                             QTextEdit, QMessageBox, QProgressBar)
from qdarkstyle import LightPalette

from lrb.common import Icon
from lrb.screenshot.lrb_screenshot import LrbScreenshot

ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("lrb_screenshot_qt")


class ScreenshotUI(QMainWindow):
    icon: QIcon = None
    w: QWidget = None
    logger: QTextEdit = None
    _thread: LrbScreenshot = None
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

        self._thread = LrbScreenshot(self)
        self._thread.value_changed.connect(lambda v: self.progress_bar.setValue(int(v * 1000)))
        self._thread.status.connect(self.status)
        self._thread.output.connect(lambda m: self.log(m))
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

        self.trans_icon()
        self.setWindowIcon(self.icon)
        self.setWindowTitle('小红书截图 - V1.0.0  Author: magese@live.cn')

    def trans_icon(self):
        data = QByteArray().fromBase64(Icon.logo.encode())
        image = QImage()
        image.loadFromData(data)
        self.icon = QIcon()
        self.icon.addPixmap(QPixmap(image), QIcon.Mode.Normal, QIcon.State.Off)

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

    def log(self, msg):
        print(msg)
        self.logger.insertPlainText(msg + '\n')

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
            QMessageBox.warning(self, '警告', '请先选择待处理excel文件及保存路径后再开始运行！')
            return

        self._thread.excel_path = self.excel_edit.text()
        self._thread.save_dir = self.save_dir_edit.text()

        self._thread.stop_status = 0
        self.excel_btn.setEnabled(False)
        self.save_dir_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.pause_btn.setEnabled(True)
        self.start_btn.setEnabled(False)

        self._thread.start()
        self.statusBar().showMessage('正在运行……')

    def pause(self):
        if self._thread.stop_status == 2:
            self.resume()
            return

        self._thread.stop_status = 2
        self.pause_btn.setEnabled(False)
        self.pause_btn.setText('暂停中...')
        self.pause_btn.setCursor(Qt.CursorShape.WaitCursor)
        self.statusBar().showMessage('暂停中……')

    def resume(self):
        self._thread.stop_status = 0
        self.pause_btn.setEnabled(False)
        self.pause_btn.setText('继续中...')
        self.pause_btn.setCursor(Qt.CursorShape.WaitCursor)
        self.statusBar().showMessage('正在运行……')

    def stop(self):
        self._thread.stop_status = 1
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
    ui.setStyleSheet(qdarkstyle.load_stylesheet(qt_api='pyqt6', palette=LightPalette()))
    ui.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
