# -*- encoding=utf8 -*-
# import sys
# from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QTextEdit
#
# class DataInputApp(QWidget):
#     def __init__(self):
#         super().__init__()
#         self.initUI()
#
#     def initUI(self):
#         self.setWindowTitle('ICD Data Input')
#         self.setGeometry(100, 100, 400, 300)
#
#
#         self.lbl_icd_name = QLabel('ICD Name:', self)
#         self.lbl_icd_name.move(20, 20)
#         self.txt_icd_name = QLineEdit(self)
#         self.txt_icd_name.move(120, 20)
#
#         self.lbl_mac_address = QLabel('MAC Address:', self)
#         self.lbl_mac_address.move(20, 60)
#
#         # 创建5个用于输入MAC地址不同部分的文本框，设置宽度和对象名称
#         text_box_width = 30
#         text_box_margin = 40
#         for i in range(1, 6):
#             txt_box = QLineEdit(self)
#             txt_box.move(120 + (i - 1) * text_box_margin, 60)
#             txt_box.setMaxLength(2)
#             txt_box.setFixedWidth(text_box_width)
#             txt_box.setObjectName(f"lineEdit_{i}")
#
#         self.btn_submit = QPushButton('Submit', self)
#         self.btn_submit.move(150, 100)
#         self.btn_submit.clicked.connect(self.writeAndReadFile)
#
#         self.txt_output = QTextEdit(self)
#         self.txt_output.setGeometry(20, 140, 360, 100)
#
#         self.show()
#
#     def writeAndReadFile(self):
#         icd_name_input = self.txt_icd_name.text()
#
#         # 获取每个文本框输入的内容
#         mac_parts = [self.findChild(QLineEdit, f"lineEdit_{i}").text() for i in range(1, 6)]
#
#         # 组合成完整的MAC地址
#         mac_address = ':'.join(mac_parts)
#
#         file_path = r'D:\SpecialTest 6\mac_address.txt'
#
#         with open(file_path, 'a') as file:
#             file.write(f"ICD{icd_name_input}- {mac_address}\n")
#
#         self.txt_icd_name.clear()
#         for i in range(1, 6):
#             self.findChild(QLineEdit, f"lineEdit_{i}").clear()
#
#         # 读取文件内容并显示在 QTextEdit 中
#         with open(file_path, 'r') as file:
#             file_content = file.read()
#             self.txt_output.setPlainText(file_content)
#
# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     ex = DataInputApp()
#     sys.exit(app.exec_())

import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QTextEdit, QFileDialog, QDesktopWidget, QMessageBox
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QProcess

# class DataInputApp(QWidget):
#     def __init__(self):
#         super().__init__()
#         self.initUI()
#
#     def initUI(self):
#         self.setWindowTitle('ICD Data Input')
#         self.setWindowIcon(QIcon(r'D:\test01\icon.ico'))
#         self.setGeometry(100, 100, 400, 300)
#         qr = self.frameGeometry()
#         cp = QDesktopWidget().availableGeometry().center()
#         qr.moveCenter(cp)
#         self.move(qr.topLeft())
#         self.lbl_icd_name = QLabel('ICD Name:', self)
#         self.lbl_icd_name.move(20, 20)
#         self.txt_icd_name = QLineEdit(self)
#         self.txt_icd_name.move(120, 20)
#
#         self.lbl_mac_address = QLabel('MAC Address:', self)
#         self.lbl_mac_address.move(20, 60)
#
#         text_box_width = 30
#         text_box_margin = 40
#         for i in range(1, 7):  # 将循环次数改为 7
#             txt_box = QLineEdit(self)
#             txt_box.move(120 + (i - 1) * text_box_margin, 60)
#             txt_box.setMaxLength(2)
#             txt_box.setFixedWidth(text_box_width)
#             txt_box.setObjectName(f"lineEdit_{i}")
#
#         self.lbl_test_engineer = QLabel('Test Engineer:', self)
#         self.lbl_test_engineer.move(20, 20)
#         self.lbl_test_engineer = QLineEdit(self)
#         self.lbl_test_engineer.move(120, 20)
#
#
#         self.btn_submit = QPushButton('Submit', self)
#         self.btn_submit.move(150, 100)
#         self.btn_submit.clicked.connect(self.writeAndReadFile)
#
#         self.btn_select_file = QPushButton('Select File', self)
#         self.btn_select_file.move(20, 100)
#         self.btn_select_file.clicked.connect(self.selectFile)
#
#         self.txt_output = QTextEdit(self)
#         self.txt_output.setGeometry(20, 140, 360, 100)
#
#
#         # self.btn_run_command = QPushButton('Run Command', self)
#         # self.btn_run_command.move(250, 100)
#         # self.btn_run_command.clicked.connect(self.runCommand)
#
#         self.show()
#
#
#     def selectFile(self):
#         options = QFileDialog.Options()
#         file_path, _ = QFileDialog.getOpenFileName(self, "Select File", "", "Text Files (*.txt)", options=options)
#         if file_path:
#             self.file_path = file_path
#
#     def writeAndReadFile(self):
#         icd_name_input = self.txt_icd_name.text()
#
#         mac_parts = [self.findChild(QLineEdit, f"lineEdit_{i}").text() for i in range(1, 7)]  # 修改为 range(1, 7)
#
#         mac_address = ':'.join(mac_parts)
#
#         with open(self.file_path, 'a') as file:
#             file.write(f"ICD{icd_name_input}- {mac_address}\n")
#
#         self.txt_icd_name.clear()
#         for i in range(1, 7):
#             self.findChild(QLineEdit, f"lineEdit_{i}").clear()
#
#         with open(self.file_path, 'r') as file:
#             file_content = file.read()
#             self.txt_output.setPlainText(file_content)
#
#     # def runCommand(self):
#     #     # 设置命令提示符的编码为 UTF-8
#     #     subprocess.run('chcp 65001', shell=True)
#     #     # 执行你的命令
#     #     cmd = 'cd /d D:\\SpecialTest 6 && python QuartServer.py'
#     #     process = subprocess.Popen(['cmd', '/c', cmd], shell=True)
#     #     process.wait()



class DataInputApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('ICD Data Input')
        self.setWindowIcon(QIcon(r'D:\test01\icon.ico'))
        self.setGeometry(100, 100, 400, 300)
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
        self.lbl_icd_name = QLabel('ICD Name:', self)
        self.lbl_icd_name.move(20, 20)
        self.txt_icd_name = QLineEdit(self)
        self.txt_icd_name.move(120, 20)

        self.lbl_mac_address = QLabel('MAC Address:', self)
        self.lbl_mac_address.move(20, 60)

        text_box_width = 30
        text_box_margin = 40
        for i in range(1, 7):
            txt_box = QLineEdit(self)
            txt_box.move(120 + (i - 1) * text_box_margin, 60)
            txt_box.setMaxLength(2)
            txt_box.setFixedWidth(text_box_width)
            txt_box.setObjectName(f"lineEdit_{i}")

        self.lbl_test_engineer = QLabel('Test Engineer:', self)
        self.lbl_test_engineer.move(20, 100)
        self.txt_test_engineer = QLineEdit(self)
        self.txt_test_engineer.move(140, 100)

        self.btn_submit = QPushButton('Submit', self)
        self.btn_submit.move(150, 140)  # Adjusted position
        self.btn_submit.clicked.connect(self.writeAndReadFile)

        self.btn_select_file = QPushButton('Select File', self)
        self.btn_select_file.move(20, 140)
        self.btn_select_file.clicked.connect(self.selectFile)

        self.txt_output = QTextEdit(self)
        self.txt_output.setGeometry(20, 180, 360, 100)  # Adjusted position

        self.show()

    def selectFile(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Select File", "", "Text Files (*.txt)", options=options)
        if file_path:
            self.file_path = file_path

    def writeAndReadFile(self):
        icd_name_input = self.txt_icd_name.text()

        mac_parts = [self.findChild(QLineEdit, f"lineEdit_{i}").text() for i in range(1, 7)]

        mac_address = ':'.join(mac_parts)

        test_engineer = self.txt_test_engineer.text()

        if not icd_name_input or not mac_address or not test_engineer:
            # 如果有一个或多个文本框为空，则弹出消息框提醒用户
            QMessageBox.warning(self, 'Warning', 'Text boxes cannot be empty.')
            return

        with open(self.file_path, 'a') as file:
            file.write(f"ICD{icd_name_input}- {mac_address}\n")

        self.txt_icd_name.clear()
        for i in range(1, 7):
            self.findChild(QLineEdit, f"lineEdit_{i}").clear()

        with open(self.file_path, 'r') as file:
            file_content = file.read()
            self.txt_output.setPlainText(file_content)

        with open('D:/test_engineer.txt', 'w') as file:
            file.write(f"Test Engineer: {test_engineer}")


if __name__ == '__main__':

    app = QApplication(sys.argv)
    ex = DataInputApp()
    sys.exit(app.exec_())
    #pyinstaller --onefile --noconsole --icon="D:\test01\icon.ico" --name=mac_address macaddress.py