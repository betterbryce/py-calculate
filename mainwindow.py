'''
Author: 895014913@qq.com 895014913@qq.com
Date: 2024-07-25 16:17:18
LastEditors: 895014913@qq.com 895014913@qq.com
LastEditTime: 2024-07-25 22:26:38
FilePath: /fof/Users/linxiaoying/Desktop/py/mainwindow.py
Description: 这是默认设置,请设置`customMade`, 打开koroFileHeader查看配置 进行设置: https://github.com/OBKoro1/koro1FileHeader/wiki/%E9%85%8D%E7%BD%AE
'''
import sys
import traceback
from PySide6.QtWidgets import QWidget, QToolTip, QPushButton, QApplication, QMessageBox, QVBoxLayout, QHBoxLayout, QLabel, QInputDialog
from PySide6.QtGui import QFont
from readExcel import setupFuc

class Ui_Form(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        QToolTip.setFont(QFont("SansSerif", 10))

        self.btn = QPushButton("生成平均单价", self)
        self.btn.setFixedSize(100, 50)
        self.btn.clicked.connect(lambda: self.handleExcel(1))

        self.btn1 = QPushButton("生成合计表", self)
        self.btn1.setFixedSize(100, 50)
        self.btn1.clicked.connect(lambda: self.handleExcel(2))

        self.label = QLabel("请点击右侧按钮计算", self)

        v_h_box = QVBoxLayout()
        v_h_box.addWidget(self.btn)
        v_h_box.addStretch(1)
        v_h_box.addWidget(self.btn1)
        v_h_box.addStretch(1)

        h_box = QHBoxLayout()
        h_box.addStretch(1)
        h_box.addWidget(self.label)
        h_box.addStretch(1)
        h_box.addLayout(v_h_box)
        h_box.addStretch(1)

        v_box = QVBoxLayout()
        v_box.addStretch(1)
        v_box.addLayout(h_box)
        v_box.addStretch(1)

        self.setLayout(v_box)
        self.resize(400, 300)
        self.setWindowTitle("计算合计表")

    def handleExcel(self, type):
        text, ok = QInputDialog.getText(self, '提示', '请输入汇率:')
        if ok and text != '':
            self.label.setText("状态：计算中")
            QApplication.processEvents()  # 强制刷新界面
            res = setupFuc(text, type)
            if res:
                self.label.setText("状态：计算完成")
            else:
                self.label.setText("请点击右侧按钮计算")
        else:
            self.label.setText("请点击右侧按钮计算")
    
    def closeEvent(self, event):
        """
        重写closeEvent, 程序结束时将stdout恢复默认
        """
        reply = QMessageBox.question(
            self, '提示', '确定要退出吗?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            sys.stdout = sys.__stdout__
            event.accept()
            res = super().closeEvent(event)
        else:
            event.ignore()
            res = None
        return res


def error_handler(exc_type, exc_value, exc_tb):
    error_message = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    reply = QMessageBox.critical(
        None, 'Error Caught!:', error_message,
        QMessageBox.StandardButton.Abort | QMessageBox.StandardButton.Retry,
        QMessageBox.StandardButton.Abort)
    if reply == QMessageBox.StandardButton.Abort:
        sys.exit(1)

def main():
    sys.excepthook = error_handler
    app = QApplication(sys.argv)
    ui = Ui_Form()
    ui.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()