
import sys
import win32com.client
from PyQt5.QtWidgets import *
from PyQt5 import QtCore, QtGui
from PyQt5.QtGui import QFont, QColor
import pandas as pd
from openpyxl import load_workbook, workbook
import datetime as dt
import copy
import os
import numpy as np
import os as oss


class SaveDialog(QMainWindow):
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.save_button_local = QPushButton('Сохранить локально', self)
        self.save_button_local.setToolTip("<h3>Сохранить локально</h3>")
        self.save_button_local.clicked.connect(self.save_table_local)
        self.save_button_local.setStyleSheet(
            "\n" "background-color: rgb(245, 193, 117);\n" "\n" "font: 14pt \"Times New Roman\";"
            "\n" "border: 2px solid rgb(96, 124, 173);\n" "border-radius: 10px;")
        self.save_button_local.setGeometry(100, 100, 300, 200)

        self.save_button_global = QPushButton('Сохранить на общий диск', self)
        self.save_button_global.setToolTip("<h3>Сохранить общий диск</h3>")
        self.save_button_global.clicked.connect(self.save_table_global)
        self.save_button_global.setStyleSheet(
            "\n" "background-color: rgb(245, 193, 117);\n" "\n" "font: 14pt \"Times New Roman\";"
            "\n" "border: 2px solid rgb(96, 124, 173);\n" "border-radius: 10px;")
        self.save_button_global.setGeometry(500, 100, 300, 200)

        self.reset_local = QPushButton('Вернуться к последней локальной \n версии', self)
        self.reset_local.setToolTip("<h3>Вернуться к последней локальной версии</h3>")
        self.reset_local.clicked.connect(self.back_local)
        self.reset_local.setStyleSheet(
            "\n" "background-color: rgb(182, 202, 237);\n" "\n" "font: 14pt \"Times New Roman\";"
            "\n" "border: 2px solid rgb(96, 124, 173);\n" "border-radius: 10px;")
        self.reset_local.setGeometry(100, 500, 300, 200)

        self.back_disk_button = QPushButton('Вернуться к общей версии', self)
        self.back_disk_button.setToolTip("<h3>Вернуться к общей версии</h3>")
        self.back_disk_button.clicked.connect(self.back_disk)
        self.back_disk_button.setStyleSheet(
            "\n" "background-color: rgb(182, 202, 237);\n" "\n" "font: 14pt \"Times New Roman\";"
            "\n" "border: 2px solid rgb(96, 124, 173);\n" "border-radius: 10px;")
        self.back_disk_button.setGeometry(500, 500, 300, 200)

        self.setWindowTitle("Сохранение данных")
        self.setGeometry(0, 0, 1000, 1000)
        # self.setStyleSheet(stylesheet12)
        # self.save_button_local.setStyleSheet("\n" "font: 15pt \"Open Sauce One SemiBold\"")
        # self.save_button_global.setStyleSheet("\n" "font: 15pt \"Open Sauce One SemiBold\"")
        # self.reset_local.setStyleSheet("\n" "font: 15pt \"Open Sauce One SemiBold\"")
        # self.back_disk_button.setStyleSheet("\n" "font: 15pt \"Open Sauce One SemiBold\"")
        self.show()


    def back_disk(self):
        # self.parent.current_filename = 'df33.xlsx'
        self.parent.current_filename = 'df2_base.xlsx'
        dff = pd.read_excel(self.parent.current_filename, sheet_name='Общая информация')
        self.parent.table1.setRows = 1
        self.parent.table1.setRows = dff.shape[0]
        # for row in range(dff.shape[0]):
        #     for col in range(dff.shape[1]):
        #         item = QTableWidgetItem(str(dff.iat[row, col]))
        #         self.parent.table1.setItem(row, col, item)

        for row in range(dff.shape[0]):
            for col in range(dff.shape[1]):
                if pd.notnull(dff.iat[row, col]):
                    if col == 1 or col == 4 or col == 7 or col == 9:
                        value = dff.iat[row, col]
                        item = QTableWidgetItem(str(value))
                    else:
                        if len(str(dff.iat[row, col])) > 0 and( (str(dff.iat[row, col]))[1].isdigit() or (str(dff.iat[row, col]))[0].isdigit()) :
                            item = QTableWidgetItem(str(int(float(dff.iat[row, col]))))
                        else:
                            item = QTableWidgetItem(dff.iat[row, col])
                    self.parent.table1.setItem(row, col, item)
                else:
                    item = QTableWidgetItem(str(''))
                    self.parent.table1.setItem(row, col, item)
        self.parent.using_file.setText("            текущий файл: " + self.parent.current_filename)

        table1 = self.parent.table1
        table2 = self.parent.table2
        table2.setRowCount(0)
        for i in range(table1.rowCount()):
            if table1.item(i, 3).text() == 'не сдал':
                table2.setRowCount(table2.rowCount() + 1)
                for j in range(table1.columnCount()):
                    table2.setItem(table2.rowCount() - 1, j, QTableWidgetItem(table1.item(i, j).text()))



    def back_local(self):
        # self.parent.current_filename = 'df33.xlsx'
        self.parent.current_filename = 'df2_local.xlsx'
        dff = pd.read_excel(self.parent.current_filename, sheet_name='Общая информация')
        self.parent.table1.setRows = 1
        self.parent.table1.setRows = dff.shape[0]
        # for row in range(dff.shape[0]):
        #     for col in range(dff.shape[1]):
        #         item = QTableWidgetItem(str(dff.iat[row, col]))
        #         self.parent.table1.setItem(row, col, item)

        for row in range(dff.shape[0]):
            for col in range(dff.shape[1]):
                if pd.notnull(dff.iat[row, col]):
                    if col == 1 or col == 4 or col == 7 or col == 9:
                        value = dff.iat[row, col]
                        item = QTableWidgetItem(str(value))
                    else:
                        if len(str(dff.iat[row, col])) > 0 and( (str(dff.iat[row, col]))[1].isdigit() or (str(dff.iat[row, col]))[0].isdigit()) :
                            item = QTableWidgetItem(str(int(float(dff.iat[row, col]))))
                        else:
                            item = QTableWidgetItem(dff.iat[row, col])
                    self.parent.table1.setItem(row, col, item)
                else:
                    item = QTableWidgetItem(str(''))
                    self.parent.table1.setItem(row, col, item)
        self.parent.using_file.setText("            текущий файл: " + self.parent.current_filename)

        table1 = self.parent.table1
        table2 = self.parent.table2
        table2.setRowCount(0)
        for i in range(table1.rowCount()):
            if table1.item(i, 3).text() == 'не сдал':
                table2.setRowCount(table2.rowCount() + 1)
                for j in range(table1.columnCount()):
                    table2.setItem(table2.rowCount() - 1, j, QTableWidgetItem(table1.item(i, j).text()))

    def save_table_local(self):
        parent = self.parent
        reply = QMessageBox.question(self, "Подтверждение",
                                     f"Вы уверены, что хотите сохранить изменения ?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        table_dict = {}
        if reply == QMessageBox.Yes:
            tabs = [parent.table1, parent.table2]
            writer1 = pd.ExcelWriter('local.xlsx')
            sheet_names = ['Общая информация', 'Должники']
            for k, tab in enumerate(tabs):
                table_dict[f"df_{k + 1}"] = pd.DataFrame(index=range(tab.rowCount()),
                                                         columns=range(tab.columnCount()))
                for col in range(tab.columnCount()):
                    for row in range(tab.rowCount()):
                        item = tab.item(row, col)
                        if item is not None:
                            table_dict[f"df_{k + 1}"].iloc[row][col] = item.text()
                df_cols = ['Отдел', 'Дата\nизменения', 'Осталось\nдней', 'Статус\nизменения', 'Дата сдачи', 'Изменение','Куратор',
                           'Дата последнего\nоповещания', \
                           'Начальник отдела', 'Дата последнего оповещания\nначальника']
                table_dict[f'df_{k + 1}'] = table_dict[f'df_{k + 1}'].set_axis(df_cols, axis=1)
                table_dict[f'df_{k + 1}'].to_excel(writer1, sheet_names[k], index=False)
            writer1.close()
            self.parent.current_filename = 'local.xlsx'
            parent.using_file.setText("            текущий файл: " + parent.current_filename)

    def save_table_global(self):

        parent = self.parent
        reply = QMessageBox.question(self, "Подтверждение",
                                     f"Вы уверены, что хотите сохранить изменения ?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        table_dict = {}
        if reply == QMessageBox.Yes:
            tabs = [parent.table1, parent.table2]
            writer1 = pd.ExcelWriter('df2_base.xlsx')
            sheet_names = ['Общая информация', 'Должники']
            for k, tab in enumerate(tabs):
                table_dict[f"df_{k + 1}"] = pd.DataFrame(index=range(tab.rowCount()),
                                                        columns=range(tab.columnCount()))
                for col in range(tab.columnCount()):
                    for row in range(tab.rowCount()):
                        item = tab.item(row, col)
                        if item is not None:
                            table_dict[f"df_{k + 1}"].iloc[row][col] = item.text()
                df_cols = ['Отдел', 'Дата\nизменения', 'Осталось\nдней', 'Статус\nизменения', 'Дата сдачи', 'Изменение', 'Куратор',\
                           'Дата последнего\nоповещания',  \
                            'Начальник отдела', 'Дата последнего оповещания\nначальника отдела']
                table_dict[f'df_{k + 1}'] = table_dict[f'df_{k + 1}'].set_axis(df_cols, axis=1)
                table_dict[f'df_{k + 1}'].to_excel(writer1, sheet_names[k], index=False)
            writer1.close()
            self.parent.current_filename = 'df2_base.xlsx'
            parent.using_file.setText("            текущий файл: " + parent.current_filename)



