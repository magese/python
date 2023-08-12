import ctypes
import datetime
import sys
from pathlib import Path

import qdarkstyle
from PyQt6.QtCore import Qt, QByteArray
from PyQt6.QtGui import QIcon, QPixmap, QTextCursor, QImage, QAction
from PyQt6.QtWidgets import (QWidget, QLineEdit, QGridLayout, QApplication, QFileDialog, QPushButton, QMainWindow,
                             QTextEdit, QMessageBox, QProgressBar, QMenuBar, QMenu, QLabel)
from qdarkstyle import LightPalette

from lrb.common import Icon
from lrb.common.Model import Lrb
from lrb.link.lrb_link import LrbLink
from lrb.negative_word.lrb_negative_word import LrbNegativeWord
from lrb.note_name.lrb_note_name import LrbNoteName
from lrb.pause.lrb_pause import LrbPause
from lrb.search_word.lrb_search_word import LrbSearchWord

ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("lrb_automatic_ui")


class LrbAutomaticUI(QMainWindow):
    icon: QIcon = None
    w: QWidget = None
    menu: QMenuBar = None
    opt: QMenu = None
    opt_text: QLabel = None
    logger: QTextEdit = None
    _thread: Lrb = None
    _link: LrbLink = None
    _name: LrbNoteName = None
    _search: LrbSearchWord = None
    _negative: LrbNegativeWord = None
    _pause: LrbPause = None
    progress_bar: QProgressBar = None
    excel_btn: QPushButton = None
    excel_edit: QLineEdit = None
    username_edit: QLineEdit = None
    password_edit: QLineEdit = None
    stop_btn: QPushButton = None
    pause_btn: QPushButton = None
    start_btn: QPushButton = None

    def __init__(self):
        super().__init__()
        self.init_ui()

    # noinspection PyUnresolvedReferences
    def create_action(self, text, shortcut, _thread):
        ac = QAction(text=text, parent=self)
        ac.setStatusTip(text)
        ac.setCheckable(True)
        ac.setChecked(False)
        ac.setShortcut('CTRL+' + shortcut)
        ac.setData(_thread)
        ac.triggered.connect(lambda b: self.change_opt(b, ac))
        self.opt.addAction(ac)

    def change_opt(self, b, ac):
        if not b:
            ac.setChecked(True)
            return
        self._thread = ac.data()
        self.opt_text.setText(ac.text())
        self.statusBar().showMessage(f'切换操作类型值：{ac.text()}')
        actions = self.opt.actions()
        for i in range(0, len(actions)):
            if b and ac != actions[i]:
                actions[i].setChecked(False)

    # noinspection PyUnresolvedReferences
    def init_ui(self):
        self.statusBar()
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setRange(0, 1000)

        self._link = LrbLink()
        self._link.value_changed.connect(lambda v: self.progress_bar.setValue(int(v * 1000)))
        self._link.status.connect(self.status)
        self._link.msg.connect(lambda m: self.log(m))
        self._link.finished.connect(self.finish)

        self._name = LrbNoteName()
        self._name.value_changed.connect(lambda v: self.progress_bar.setValue(int(v * 1000)))
        self._name.status.connect(self.status)
        self._name.msg.connect(lambda m: self.log(m))
        self._name.finished.connect(self.finish)

        self._search = LrbSearchWord()
        self._search.value_changed.connect(lambda v: self.progress_bar.setValue(int(v * 1000)))
        self._search.status.connect(self.status)
        self._search.msg.connect(lambda m: self.log(m))
        self._search.finished.connect(self.finish)

        self._negative = LrbNegativeWord()
        self._negative.value_changed.connect(lambda v: self.progress_bar.setValue(int(v * 1000)))
        self._negative.status.connect(self.status)
        self._negative.msg.connect(lambda m: self.log(m))
        self._negative.finished.connect(self.finish)

        self._pause = LrbPause()
        self._pause.value_changed.connect(lambda v: self.progress_bar.setValue(int(v * 1000)))
        self._pause.status.connect(self.status)
        self._pause.msg.connect(lambda m: self.log(m))
        self._pause.finished.connect(self.finish)

        w = QWidget()
        self.setCentralWidget(w)

        self.menu = self.menuBar()
        self.opt = self.menu.addMenu('操作')
        self.opt_text = QLabel()
        self.create_action('更换监测链接', 'L', self._link)
        self.create_action('更换笔记名称', 'N', self._name)
        self.create_action('更换搜索词', 'S', self._search)
        self.create_action('添加否定词', 'F', self._negative)
        self.create_action('暂停笔记', 'P', self._pause)
        ac = self.opt.actions()[0]
        ac.setChecked(True)
        self._thread = self._link
        self.opt_text.setText(ac.text())

        opt = QLabel('操作类型')
        exl = QLabel('Excel')
        un = QLabel('用户名')
        pwd = QLabel('密码')

        self.username_edit = QLineEdit(self)
        self.password_edit = QLineEdit(self)
        self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)

        self.excel_btn = QPushButton('选择excel')
        self.excel_edit = QLineEdit(self)
        self.excel_edit.setReadOnly(True)
        self.excel_btn.clicked.connect(self.excel_select)

        self.logger = QTextEdit(self)
        self.logger.setReadOnly(True)
        self.logger.setMinimumHeight(200)
        self.logger.setTextInteractionFlags(Qt.TextInteractionFlag.NoTextInteraction)
        self.logger.textChanged.connect(self.scroll_to_end)
        self.logger.viewport().setCursor(Qt.CursorShape.ArrowCursor)

        self.stop_btn = QPushButton('停止运行')
        self.pause_btn = QPushButton('暂停运行')
        self.start_btn = QPushButton('开始运行')
        clear_btn = QPushButton('清空日志')
        save_btn = QPushButton('保存日志')

        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop)
        self.pause_btn.setEnabled(False)
        self.pause_btn.clicked.connect(self.pause)
        self.start_btn.clicked.connect(self.start)
        clear_btn.clicked.connect(lambda: self.logger.clear())
        save_btn.clicked.connect(self.save_log)

        grid = QGridLayout()
        grid.setSpacing(10)

        grid.addWidget(opt, 1, 0)
        grid.addWidget(self.opt_text, 1, 1, 1, 7)

        grid.addWidget(exl, 2, 0)
        grid.addWidget(self.excel_edit, 2, 1, 1, 7)
        grid.addWidget(self.excel_btn, 2, 7)

        grid.addWidget(un, 3, 0)
        grid.addWidget(self.username_edit, 3, 1, 1, 7)

        grid.addWidget(pwd, 4, 0)
        grid.addWidget(self.password_edit, 4, 1, 1, 7)

        grid.addWidget(clear_btn, 5, 0)
        grid.addWidget(save_btn, 5, 1)
        grid.addWidget(self.stop_btn, 5, 5)
        grid.addWidget(self.pause_btn, 5, 6)
        grid.addWidget(self.start_btn, 5, 7)

        grid.addWidget(self.progress_bar, 6, 0, 1, 8)

        grid.addWidget(self.logger, 7, 0, 8, 8)

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
        msg = f'选择待处理文件：{fname[0]}'
        self.statusBar().showMessage(msg)

    def save_log(self):
        now = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        filename = f'{self.opt_text.text()}-{now}.txt'
        home_dir = f'{str(Path.home())}\\Desktop\\{filename}'
        filepath, _ = QFileDialog.getSaveFileName(self.centralWidget(), '选择保存的路径', home_dir, '文件类型 (*.txt)')
        print(filepath)
        if str(filepath) == '':
            QMessageBox.information(self, "提示", "请选择要保存的文件路径")
        else:
            with open(filepath, "w") as f:
                f.write(self.logger.toPlainText())
            self.statusBar().showMessage(f'日志文件保存成功：{filepath}')

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
        if len(excel_text) == 0:
            QMessageBox.warning(self, '警告', '请先选择待处理excel文件及保存路径后再开始运行！')
            return

        self._thread._excel_path = self.excel_edit.text()
        self._thread._username = self.username_edit.text()
        self._thread._password = self.password_edit.text()

        self._thread.stop_status = 0
        self.excel_edit.setEnabled(False)
        self.username_edit.setEnabled(False)
        self.password_edit.setEnabled(False)
        self.excel_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.pause_btn.setEnabled(True)
        self.start_btn.setEnabled(False)
        for ac in self.opt.actions():
            ac.setEnabled(False)

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
        self.excel_edit.setEnabled(True)
        self.username_edit.setEnabled(True)
        self.password_edit.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.stop_btn.setText('停止中...')
        self.stop_btn.setCursor(Qt.CursorShape.WaitCursor)
        for ac in self.opt.actions():
            ac.setEnabled(True)
        self.statusBar().showMessage('停止中……')

    def finish(self):
        self.excel_edit.setEnabled(True)
        self.username_edit.setEnabled(True)
        self.password_edit.setEnabled(True)
        self.excel_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.pause_btn.setEnabled(False)
        self.start_btn.setEnabled(True)
        self.progress_bar.reset()
        for ac in self.opt.actions():
            ac.setEnabled(True)
        self.statusBar().showMessage('已完成')


def main():
    app = QApplication(sys.argv)
    ui = LrbAutomaticUI()
    ui.setStyleSheet(qdarkstyle.load_stylesheet(qt_api='pyqt6', palette=LightPalette()))
    ui.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
