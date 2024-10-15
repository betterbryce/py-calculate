# -*- coding: utf-8 -*-、
import logging
import inspect
import traceback
import os
import sys
from datetime import datetime
from random import randint
from PyQt6.QtCore import pyqtSignal, QObject, QtMsgType, qInstallMessageHandler
from PyQt6.QtWidgets import QDialog, QApplication, QTextEdit, QPushButton, QMessageBox
 
 
class Stream(QObject):
    """
    重定向控制台输出到文本框控件
    """
    newText = pyqtSignal(str)
 
    # 任何定义了类似于文件write方法的对象可以指定给sys.stdout,
    # 所有的标准输出将发送到该方法对象上
    def write(self, text):
        self.newText.emit(str(text))
        QApplication.processEvents()
 
 
class QTextEditHandler(logging.Handler):
    def __init__(self, parent):
        super().__init__()
        self.text_edit = parent
        self.formatter = logging.Formatter('%(asctime)s - [%(filename)s][line:%(lineno)d][%(levelname)s]  %(message)s')
 
    def emit(self, record):
        msg = self.format(record)
        self.text_edit.append(msg)
        self.text_edit.ensureCursorVisible()
 
    def format(self, record):
        if record.levelno == logging.DEBUG:
            color = 'gray'
        elif record.levelno == logging.INFO:
            color = 'black'
        elif record.levelno == logging.WARNING:
            color = 'orange'
        elif record.levelno == logging.ERROR:
            color = 'darkRed'
        elif record.levelno == logging.CRITICAL or record.levelno == logging.FATAL:
            color = 'red'
        else:
            color = 'blue'
        new_msg = self.formatter.format(record)
        msg = '<span style="color:{}">{}</span>'.format(color, new_msg)
        return msg
 
 
class Ui_Form(QDialog):
    """
    主界面
    """
 
    def __init__(self):
        super().__init__()
        self.resize(1000, 500)
        self.textEdit = QTextEdit(self)
        self.textEdit.resize(900, 200)
        self.QPushButton1 = QPushButton('点我运行正常程序', self)
        self.QPushButton1.clicked.connect(self.normal_func)
        self.QPushButton1.move(10, 250)
        self.QPushButton2 = QPushButton('点我运行异常程序', self)
        self.QPushButton2.clicked.connect(self.error_func)
        self.QPushButton2.move(10, 280)
 
        # 自定义输出流
        sys.stdout = Stream()
        sys.stdout.newText.connect(self.onUpdateText)
 
        # Log打印
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.DEBUG)
        formatter = logging.Formatter('%(asctime)s - [%(filename)s][line:%(lineno)d][%(levelname)s]  %(message)s')
        # QTextEdit-Handler
        handler = QTextEditHandler(self.textEdit)
        handler.setFormatter(formatter)
        self.logger.addHandler(handler)
        # File-Handler
        log_file_name = os.path.join(os.path.dirname(__file__), 'Log_Folder', 'log_{time}_{randomn}'.format(
            time=datetime.now().strftime('%Y_%m%d_%H%M%S'), randomn=randint(0, 9)))
        os.makedirs(os.path.dirname(log_file_name), exist_ok=True)
        file_handler = logging.FileHandler(log_file_name)
        file_handler.setFormatter(formatter)
        self.logger.addHandler(file_handler)
 
        # QT 消息处理
        qInstallMessageHandler(self.redirection_msg)
 
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
 
    def onUpdateText(self, text):
        """
        重定向控制台输出到文本框控件
        """
        if text == '\n':
            # print函数默认会在输出内容后自动添加一个换行符\n
            # 后续脚本能够保证，每次显示print内容时，都是从一个新行的行首开始，不再需要这个换行符，跳过即可
            return
        cursor = self.textEdit.textCursor()
 
        if cursor.position() == cursor.block().position():
            # 如果当前光标位置是行首, 直接打印文字内容
            cursor.insertText(text)
        else:
            # 如果当前光标位置是行中, 换行，从下一行行首开始打印文字内容
            cursor.movePosition(cursor.MoveOperation.End)
            cursor.insertText('\n' + text)
        self.textEdit.setTextCursor(cursor)
        self.textEdit.ensureCursorVisible()
        return
 
    def error_func(self):
        """
        执行后会报错的程序
        """
        self.logger.info('运行存在异常的程序脚本')
        a = [1, 2, 3, 4]
        self.logger.info(f'程序运行结果: {a[5]}')
 
    def normal_func(self):
        """
        正常运行的程序
        """
        self.logger.info('INFO级别：用于展示程序的运行状态和正常操作')
        self.logger.debug('DEBUG级别：用于输出详细的运行日志')
        self.logger.warning(
            'WARNING级别：用于记录一些警告信息，这些信息可能表明程序出现了一些异常或潜在的问题，但并不会导致程序的崩溃或功能的异常')
        self.logger.error('ERROR级别：用于记录程序中的错误信息，但仍不影响系统的继续运行')
        self.logger.critical('CRITICAL级别：最高级别的日志，用于记录知名错误，这些错误会导致程序立即退出，并且无法恢复；')
        self.logger.fatal('FATAL级别：与CRITICAL一样表示最高级别的日志，在不同的框架或库中，可能会定义不同的日志级别，'
                          'Python标准日志模块中没有定义FATAL级别，某些第三方日志库可能会使用FATAL级别表示同样的错误')
        a = [1, 2, 3, 4]
        self.logger.info(f'程序运行结果: {a[2]}')
        print('这是第一段print文字，默认从新的一行行首开始显示')
        print('这是第二段print文字，默认从新的一行行首开始显示')
        print('这是第三段print文字')
 
    def redirection_msg(self, mode, context, message):
        """
        打印错误信息并且弹出警告窗口
        """
        frame = inspect.currentframe().f_back
        filename = frame.f_code.co_filename
        lineno = frame.f_lineno
        funcname = frame.f_code.co_name
        if mode == QtMsgType.QtInfoMsg:
            mode = 20
        elif mode == QtMsgType.QtWarningMsg:
            mode = 30
        elif mode == QtMsgType.QtCriticalMsg:
            mode = 50
        elif mode == QtMsgType.QtFatalMsg:
            mode = 50
        else:
            mode = 10
 
        record = self.logger.makeRecord(name=self.logger.name, level=mode, fn=filename, lno=lineno, msg=message, args=(), exc_info=None, func=funcname, sinfo=None)
        self.logger.handle(record)
 
 
def error_handler(exc_type, exc_value, exc_tb):
    error_message = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    reply = QMessageBox.critical(
        None, 'Error Caught!:', error_message,
        QMessageBox.StandardButton.Abort | QMessageBox.StandardButton.Retry,
        QMessageBox.StandardButton.Abort)
    if reply == QMessageBox.StandardButton.Abort:
        sys.exit(1)
 
 
if __name__ == '__main__':
    sys.excepthook = error_handler
    app = QApplication(sys.argv)
    ui = Ui_Form()
    ui.show()
    sys.exit(app.exec())
 