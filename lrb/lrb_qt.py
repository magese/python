import sys

from PyQt6.QtWidgets import QApplication, QWidget, QPushButton


def main():
    app = QApplication(sys.argv)
    w = QWidget()
    w.setWindowTitle('Simple')
    btn = QPushButton('Hello PyQt6!', w)
    btn.move(50, 50)
    w.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
