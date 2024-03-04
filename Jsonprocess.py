# -*- encoding=utf8 -*-
import json
import os
import sys
from datetime import datetime

import pandas as pd
from PyQt5.QtCore import QTranslator
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtWidgets import QFileDialog
from openpyxl import Workbook
from openpyxl.styles import Font

import json_process_ui


class MainDialog(QMainWindow, json_process_ui.Ui_MainWindow):
    def __init__(self):
        super(MainDialog, self).__init__()
        self.setupUi(self)
        self.trans = QTranslator()
        _app = QApplication.instance()
        self.setWindowTitle("测试数据处理脚本V1.0")

        self.pushButton.clicked.connect(self.showFileDialog)
        self.pushButton_2.clicked.connect(self.json_excel)
        self.pushButton_3.clicked.connect(self.json_set)

    def showFileDialog(self):
        # 打开文件对话框
        file_dialog = QFileDialog()
        file_dialog.exec_()
        # 获取用户选择的文件路径
        file_paths = file_dialog.selectedFiles()
        if file_paths:
            self.textEdit.clear()
            for file_path in file_paths:
                self.textEdit.append(file_path)

    def getFilePaths(self):
        file_paths = self.textEdit.toPlainText().split('\n')
        return file_paths

    def json_excel(self):
        # 读取JSON文件
        file_path = self.textEdit.toPlainText()
        if file_path:
          with open(file_path) as file:
            data = json.load(file)
        wb = Workbook()
        ws = wb.active
        # 将JSON数据转换为DataFrame对象

        df = pd.DataFrame.from_dict(data.items())
        df.columns = ['测试项目', 'Result']
        ws.append(df.columns.tolist())

        for row in df.iterrows():
            ws.append(row[1].tolist())

        for cell in ws['B']:
            if cell.value == 'fail':
                cell.font = Font(color="FF0000")
            elif cell.value == 'pass':
                cell.font = Font(color="0000FF")
        now = datetime.now()
        timestamp = now.strftime("%Y-%m-%d_%H-%M-%S")
        excel_file_name = f"{timestamp}.xlsx"
        desktop_path = os.path.expanduser("~/Desktop")
        excel_file_path = os.path.join(desktop_path, excel_file_name)

        # 保存Excel文件
        wb.save(excel_file_path)

    def json_set(self):
        file_path = self.textEdit.toPlainText()
        if file_path:
            with open(file_path) as file:
                data = json.load(file)
            formatted_data = json.dumps(data, indent=4)

            now = datetime.now()
            timestamp = now.strftime("%Y-%m-%d_%H-%M-%S")
            json_file_name = f"{timestamp}.json"
            desktop_path = os.path.expanduser("~/Desktop")
            excel_file_path = os.path.join(desktop_path, json_file_name)

            # 将格式化后的JSON数据写入新的JSON文件
            with open(excel_file_path , 'w') as file:
                file.write(formatted_data)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = MainDialog()
    win.show()
    sys.exit(app.exec_())