import sys
import warnings
import datetime as dt
import numpy as np
import pandas as pd
from PyQt5 import Qt, QtCore
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import QMainWindow, QPushButton, QLabel, QLineEdit, QMessageBox, QTabWidget, QTableWidget, \
    QVBoxLayout, QWidget, QHBoxLayout, QTableWidgetItem, QCheckBox, QComboBox, QDateEdit, QCalendarWidget, QApplication
import xlwings as xw
from mail_module import head_dep_message, kurator_message, send_message_heads

warnings.filterwarnings('ignore')
import shutil
import time
from utils import log_print

try:
    with open('путь.txt', 'r', encoding="utf-8") as file:
        put = file.readline().rstrip('\n')
    src_file = f'{put}'
    dst_folder = 'kadrifile.xlsx'
    shutil.copy(src_file, dst_folder)

except Exception as e:
    log_print('Не удается получить доступ к облачному диску', e)
    try:
        with open('путь.txt', 'r', encoding="utf-8") as file:
            file.readline().rstrip('\n')
            put = file.readline().rstrip('\n')
        src_file = f'{put}'
        dst_folder = 'kadrifile.xlsx'
        shutil.copy(src_file, dst_folder)
    except Exception as e:
        log_print('Не удается получить доступ к облачному диску', e)
        pass



app = xw.App(visible=False)
wb = xw.Book('kadrifile.xlsx')
wb.save('kadrifile.xlsx')
wb.close()
app.quit()


class App1(QMainWindow):
    def __init__(self):
        super().__init__()
        self.title = 'Авторизация отдел кадров'

        self.pushButton = QPushButton("ВОЙТИ", self)
        self.pushButton.setGeometry(QtCore.QRect(843, 737, 253, 58))
        self.pushButton.setToolTip("<h3>Пройти верификацию</h3>")
        self.pushButton.clicked.connect(self.cheklogpas)
        self.pushButton.setStyleSheet("background-color: rgb(33, 53, 89);\n"
                                      "color: white;\n"
                                      "font: 16pt Myriad pro;\n"
                                      "font-weight: bold;\n"
                                      "\n" "border: 0px solid rgb(6, 73, 129);\n" "border-radius: 15px;")

        self.label2 = QLabel("Введите логин", self)
        self.label2.setGeometry(QtCore.QRect(745, 438, 450, 50))
        self.label2.setAcceptDrops(True)
        self.label2.setAutoFillBackground(False)
        self.label2.setScaledContents(True)
        # self.label2.setAlignment(QtCore.Qt.AlignCenter)
        self.label2.setWordWrap(True)
        self.label2.setStyleSheet("\n" "color: rgb(33, 53, 89);\n"
                                  "\n" "background-color: rgb(255, 255, 255,0);\n" "\n" "font: 16pt \"Myriad pro\";\n" "\n"
                                  "font-weight: bold")

        self.label3 = QLabel("Введите пароль", self)
        self.label3.setGeometry(QtCore.QRect(745, 578, 450, 50))
        self.label3.setAcceptDrops(True)
        self.label3.setAutoFillBackground(False)
        self.label3.setScaledContents(True)
        # self.label3.setAlignment(QtCore.Qt.AlignCenter)
        self.label3.setWordWrap(True)
        self.label3.setStyleSheet("\n" "color: rgb(33, 53, 89);\n"
                                  "\n" "background-color: rgb(255, 255, 255,0);\n" "\n" "font: 16pt \"Myriad pro\";\n" "\n"
                                  "font-weight: bold")

        self.text1 = '0'
        self.lineEdit = QLineEdit(self)
        self.lineEdit.setGeometry(QtCore.QRect(745, 485, 442, 55))
        self.lineEdit.setObjectName("<h3>Start the Session</h3>")
        self.lineEdit.setAlignment(QtCore.Qt.AlignCenter)
        self.lineEdit.setStyleSheet(
            "\n" "background-color: rgb(239, 239, 238);\n" "\n" "font: 18pt \"Times New Roman\";"
            "\n" "border: 0px solid rgb(217, 217, 217);\n" "border-radius: 8px;")

        self.text2 = '0'
        self.lineEdit2 = QLineEdit(self)
        self.lineEdit2.setGeometry(QtCore.QRect(745, 625, 442, 55))
        self.lineEdit2.setObjectName("<h3>Start the Session</h3>")
        self.lineEdit2.setAlignment(QtCore.Qt.AlignCenter)
        self.lineEdit2.setStyleSheet(
            "\n" "background-color: rgb(239, 239, 238);\n" "\n" "font: 18pt \"Times New Roman\";"
            "\n" "border: 0px solid rgb(217, 217, 217);\n" "border-radius: 8px;")

        self.text3 = '0'
        self.lineEdit3 = QLineEdit(self)
        self.lineEdit3.setGeometry(QtCore.QRect(580, 280, 730, 50))
        self.lineEdit3.setObjectName("<h3>Start the Session</h3>")
        self.lineEdit3.setAlignment(QtCore.Qt.AlignCenter)
        self.lineEdit3.setStyleSheet(
            "\n" "background-color: rgb(217, 217, 217);\n" "\n" "font: 18pt \"Times New Roman\";"
            "\n" "border: 0px solid rgb(217, 217, 217);\n" "border-radius: 10px;")
        self.lineEdit3.setVisible(False)

        self.pushButton2 = QPushButton("", self)
        self.pushButton2.setGeometry(QtCore.QRect(1385, 800, 80, 80))
        self.pushButton2.clicked.connect(self.opendoor)
        self.pushButton2.setStyleSheet(
            "\n" "background-color: rgba(0, 0, 0, 0);\n" "\n" "font: 10pt \"Times New Roman\";"
            "\n" "border-radius: 12px;")

        self.pushButton7 = QPushButton("Поменять путь выгрузки файлов", self)
        self.pushButton7.setGeometry(QtCore.QRect(580, 200, 700, 50))
        self.pushButton7.clicked.connect(self.change_opendoor)
        self.pushButton7.setStyleSheet(
            "\n" "background-color: rgb(127, 155, 205);\n" "\n" "font: 18pt \"Times New Roman\";"
            "\n" "border: 3px solid rgb(6, 73, 129);\n" "border-radius: 30px;")
        self.pushButton7.setVisible(False)
        self.showMaximized()
        self.show()

    def opendoor(self):
        print('gggg')
        self.text1 = self.lineEdit.text()
        self.text2 = self.lineEdit2.text()
        self.loginn1 = 'admin'
        self.loginn2 = 'admin'
        if self.loginn1 == self.text1 and self.loginn2 == self.text2:
            self.lineEdit3.setVisible(True)
            self.pushButton7.setVisible(True)

    def change_opendoor(self):
        self.putfiles = self.lineEdit3.text()
        try:
            with open('путь.txt', 'w') as file:
                file.truncate(0)
            with open("путь.txt", "w", encoding="utf-8") as file:
                file.write(self.putfiles)

        except FileNotFoundError:
            QMessageBox.warning(self, "Ошибка777", "Файл проверки лог/пас не найден.", '9')

    def cheklogpas(self):
        global glav
        glav = 0
        self.text1 = self.lineEdit.text()
        print(self.text1)
        self.text2 = self.lineEdit2.text()
        self.r = 0
        self.loginn1 = 'admin'
        self.loginn2 = 'admin'
        try:
            with open("логин.txt", "r", encoding="utf-8") as file:
                for line in file:
                    # Отделяем первый набор символов до знака "№"
                    login = line.split("/")[0]
                    password = line.split("/")[1]
                    global otdel
                    otdel = line.split("/")[2]
                    # Сравниваем с поисковым запросом
                    if login == self.text1 and password == self.text2:

                        # Если найдено соответствие, выводим строчку
                        QMessageBox.information(self, "Найдено соответствие", line)
                        if self.loginn1 == self.text1 and self.loginn2 == self.text2:
                            glav = 1
                        self.otkrit()
                        self.r = 1
                        break

                if self.r == 0:
                    QMessageBox.warning(self, "Ошибка 333", 'Ошибка вы ввели неверный логин или пароль')
                else:
                    pass
        except FileNotFoundError:
            QMessageBox.warning(self, "Ошибка777", "Файл проверки лог/пас не найден.", '9')

    def otkrit(self):
        try:
            global proverka
            proverka = 0
            self.w = MyWindow()
            app.setStyleSheet(stylesheet12)
            self.w.showMaximized()
            self.w.show()
            self.hide()
        except Exception as e:
            log_print(e)


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.current_filename = 'df2_base.xlsx'
        self.df_general_information = pd.read_excel(self.current_filename, sheet_name='Общая информация')
        order = self.df_general_information.columns
        heads = self.get_heads_and_kurators()
        heads = heads.rename(columns={'Структурное подразделение': 'Отдел', 'Начальник отдела/Первый по должности если он отсутствует': 'Начальник отдела'})
        self.df_general_information = self.df_general_information.merge(heads, on='Отдел', how='right')
        self.df_general_information.rename(columns={'Куратор_y' : 'Куратор', 'Начальник отдела_y' : 'Начальник отдела'}, inplace=True)
        self.df_general_information.drop(columns=['Куратор_x','Начальник отдела_x'], inplace=True)
        self.df_general_information = self.df_general_information[order]

        self.departments = pd.read_excel('kadrifile.xlsx', sheet_name='Штатка', skiprows=0, usecols='B')
        self.departments = self.departments.drop_duplicates()['Подразделение'].tolist()

        self.window_list = list()

        # Создаём QTabWidget
        self.tabs = QTabWidget()

        self.setWindowTitle("Отдел внутреннего контроля")
        self.setGeometry(0, 0, 1920, 1080)

        # Создаем QTableWidget
        self.general_information_table = QTableWidget()
        self.general_information_table.setColumnCount(len(self.df_general_information.columns))
        self.general_information_table.setHorizontalHeaderLabels(self.df_general_information.columns)


        # self.Debts_table = QTableWidget()
        # self.Debts_table.setColumnCount(len(self.df_debts.columns))
        # self.Debts_table.setHorizontalHeaderLabels(self.df_debts.columns)
        # self.fill_table_debts()


        self.kvk_list = KVK_list_tab()

        #Добавляем табы
        self.tabs.addTab(self.general_information_table, "Общая информация")
        self.tabs.addTab(self.kvk_list, "Список КВК")


        self._fill_general_dataframe()
        self.fill_table_general_information()

        layout = QVBoxLayout()
        layout.setContentsMargins(20, 60, 20, 30)

        widget = QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)

        self.nach_button = QPushButton('Рассылка начальникам отделов', self)
        # self.staff_button.setToolTip("<h3>Пройти верификацию и разослать уведомление сотрудникам</h3>")
        self.nach_button.clicked.connect(self.message_nach)
        self.nach_button.setStyleSheet(
            "\n" "background-color: rgb(245, 193, 117);\n" "\n" "font: 14pt \"Times New Roman\";"
            "\n" "border: 2px solid rgb(96, 124, 173);\n" "border-radius: 10px;")
        self.nach_button.setFixedHeight(35)

        self.current_nach_button = QPushButton('Точечная рассылка начальнику', self)
        # self.staff_button.setToolTip("<h3>Пройти верификацию и разослать уведомление сотрудникам</h3>")
        self.current_nach_button.clicked.connect(self.message_current_nach)
        self.current_nach_button.setStyleSheet(
            "\n" "background-color: rgb(245, 193, 117);\n" "\n" "font: 14pt \"Times New Roman\";"
            "\n" "border: 2px solid rgb(96, 124, 173);\n" "border-radius: 10px;")
        self.current_nach_button.setFixedHeight(35)

        self.kurator_button = QPushButton('Рассылка кураторам отделов', self)
        self.kurator_button.clicked.connect(self.message_kurator)
        self.kurator_button.setStyleSheet(
            "\n" "background-color: rgb(245, 193, 117);\n" "\n" "font: 14pt \"Times New Roman\";"
            "\n" "border: 2px solid rgb(96, 124, 173);\n" "border-radius: 10px;")
        self.kurator_button.setFixedHeight(35)

        self.save_button = QPushButton('Сохранить', self)
        self.save_button.clicked.connect(self.saving)
        self.save_button.setStyleSheet(
            "\n" "background-color: rgb(182, 202, 237);\n" "\n" "font: 14pt \"Times New Roman\";"
            "\n" "border: 2px solid rgb(96, 124, 173);\n" "border-radius: 10px;")
        self.save_button.setFixedHeight(35)


        self.add_new_KVK = QPushButton('Добавить новую карту контроля')
        self.add_new_KVK.clicked.connect(self.kvk_list.add_new_card)
        self.add_new_KVK.setStyleSheet(
            "\n" "background-color: rgb(182, 202, 237);\n" "\n" "font: 14pt \"Times New Roman\";"
            "\n" "border: 2px solid rgb(96, 124, 173);\n" "border-radius: 10px;")
        self.add_new_KVK.setFixedHeight(35)

        self.del_last_KVK = QPushButton('Удалить последнюю карту контроля')
        self.del_last_KVK.clicked.connect(self.kvk_list.del_last_card)
        self.del_last_KVK.setStyleSheet(
            "\n" "background-color: rgb(182, 202, 237);\n" "\n" "font: 14pt \"Times New Roman\";"
            "\n" "border: 2px solid rgb(96, 124, 173);\n" "border-radius: 10px;")
        self.del_last_KVK.setFixedHeight(35)



        horizontal_layout = QHBoxLayout()
        horizontal_layout.addWidget(self.nach_button)
        horizontal_layout.addWidget(self.current_nach_button)
        horizontal_layout.addWidget(self.kurator_button)
        horizontal_layout.addWidget(self.save_button)
        horizontal_layout.addWidget(self.add_new_KVK)
        horizontal_layout.addWidget(self.del_last_KVK)

        layout.addSpacing(15)
        layout.addLayout(horizontal_layout)
        layout.addSpacing(20)
        layout.addWidget(self.tabs)
        layout.addSpacing(15)
        self.current_tab_index = 0
        self.tabs.currentChanged.connect(self.tab_changed)
        self.showMaximized()

    def _fill_general_dataframe(self):
        """Заполнить пандасовский датафрейм. в соответствии с 3й вкладкой
        serias имеет отдел и данные
        """
        self.df_general_information.drop(columns=['Осталось дней', 'Количество карт контроля', 'Срок сдачи'],
                                         inplace=True)
        self.df_general_information.insert(1, 'Количество карт контроля', np.nan)
        self.df_general_information.insert(1, 'Осталось дней', np.nan)
        self.df_general_information.insert(1, 'Срок сдачи', np.nan)

        cards_preproc = self.kvk_list.card_inner_control[self.kvk_list.card_inner_control['Статус'].isna()]
        count_cards = cards_preproc.groupby('Отдел').count()['Осталось дней'].reset_index()
        deadline_frame = 1

        self.df_general_information = self.df_general_information

        tmp = self.df_general_information.merge(count_cards, on=['Отдел'], how='left')
        self.df_general_information['Количество карт контроля'] = tmp['Осталось дней_y']
        self.df_general_information['Количество карт контроля'] = self.df_general_information['Количество карт контроля'].astype('Int64')

        for dep in cards_preproc['Отдел'].drop_duplicates():
            tmp = cards_preproc[cards_preproc['Отдел'] == dep].iloc[0]
            self.df_general_information.loc[self.df_general_information['Отдел'] == dep, 'Срок сдачи'] = tmp['Срок сдачи']
            self.df_general_information.loc[self.df_general_information['Отдел'] == dep, 'Осталось дней'] = tmp['Осталось дней']
        self.df_general_information['Осталось дней'] = self.df_general_information['Осталось дней'].astype('Int64')

    def fill_table_general_information(self):
        """
        Заполняем таблицу с общей информацией
        """
        self._fill_general_dataframe()
        self.df_general_information = self.df_general_information
        self.departments = self.departments
        self.departments = self.departments
        self.general_information_table = self.general_information_table

        self.general_information_table.setRowCount(0)
        for i, row in self.df_general_information.iterrows():
            self.general_information_table.setRowCount(self.general_information_table.rowCount() + 1)
            for icol in range(self.general_information_table.columnCount()):
                value = str(row[icol])
                if value != 'nan' and value != '<NA>' and value != np.nan:
                    if icol == 2 and row[icol] <= 0:
                        self.general_information_table.setItem(i, icol, QTableWidgetItem(value))
                        self.general_information_table.item(i, icol).setBackground(QColor(238, 216, 221))
                    else:
                        self.general_information_table.setItem(i, icol, QTableWidgetItem(value))
                else:
                    self.general_information_table.setItem(i, icol, QTableWidgetItem(''))
        self.general_information_table.resizeColumnsToContents()

    def get_heads_and_kurators(self):
        df_heads_kurs = pd.read_excel('kadrifile.xlsx', sheet_name='Историярасслылок2', skiprows=3, usecols='B:D')
        return df_heads_kurs

    def tab_changed(self, index):
        self.fill_table_general_information()

        # Сохраняем индекс текущей вкладки
        self.current_tab_index = index

    def set_row_color(self):
        if self.tabs.currentWidget() == self.general_information_table:
            self._color_general_information()

    def on_combobox_changed(self, text):
        """Считает просрачку, когда тыкаешь несдал"""
        self.general_information_table.currentItem().setText(text)
        if text == 'не сдал':
            row = self.general_information_table.currentItem().row()
            new_text = str(14 - (dt.datetime.now() - pd.to_datetime(self.general_information_table.item(row, 1).text(),
                                                                    format='%d.%m.%Y')) // np.timedelta64(1, 'D'))
            self.general_information_table.item(row, 2).setText(new_text)

    def saving(self):
        # new_window = SaveDialog(self)
        # self.window_list.append(new_window)

        reply = QMessageBox.question(self, "Подтверждение",
                                     f"Вы уверены, что хотите сохранить изменения ?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        table_dict = {}
        if reply == QMessageBox.Yes:
            tabs = [self.general_information_table]
            writer1 = pd.ExcelWriter('df2_base.xlsx')
            sheet_names = ['Общая информация']
            for k, tab in enumerate(tabs):
                table_dict[f"df_{k + 1}"] = pd.DataFrame(index=range(tab.rowCount()),
                                                         columns=range(tab.columnCount()))
                for col in range(tab.columnCount()):
                    for row in range(tab.rowCount()):
                        item = tab.item(row, col)
                        if item is not None:
                            table_dict[f"df_{k + 1}"].iloc[row][col] = item.text()

                df_cols = ['Отдел',	'Срок сдачи', 'Осталось дней', 'Количество карт контроля', 'Куратор', 'Дата последнего оповещения', 'Начальник отдела', 'Дата последнего оповещения начальникам отдела']

                table_dict[f'df_{k + 1}'] = table_dict[f'df_{k + 1}'].set_axis(df_cols, axis=1)
                table_dict[f'df_{k + 1}'].to_excel(writer1, sheet_names[k], index=False)

            self.kvk_list.card_inner_control.to_excel(writer1, 'Карты внутреннего контроля', index=False)


            writer1.close()

    def message_nach(self):
        try:
            head_dep_message(self.df_general_information)
            self.df_general_information.loc[self.df_general_information['Осталось дней'].notna(), 'Дата последнего оповещения начальникам отдела'] = dt.datetime.now().strftime('%d.%m.%Y')
            self.fill_table_general_information()
            QMessageBox.about(self, "Рассылка начальникам отдела", "Рассылка отправлена всем начальникам отделов")

        except Exception as e:
            QMessageBox.about(self, "Рассылка начальникам отдела", "Рассылка не отправлена непредвиденная ошибка, обратитесь к администраторам.")
            log_print('Сообщения упали', e)

    def message_current_nach(self):
        current_row = self.general_information_table.currentRow()
        if current_row == -1:
            QMessageBox.about(self, "Рассылка начальнику отдела", "Не выбран конкретный отдел")
            log_print('Не выбран начальник для точечного сообщения')
            return -1
        try:
            df_names_emails = pd.read_excel('kadrifile.xlsx', sheet_name='Историярасслылок2', skiprows=3, usecols='J:M')

            df_current_row = self.df_general_information.iloc[current_row]
            name = df_current_row['Начальник отдела']
            department = df_current_row['Отдел']
            email = df_names_emails[df_names_emails['Отдел'] == department]['Почта.1'].to_string(index=False)
            deadline = df_current_row['Срок сдачи']
            left_days = df_current_row['Осталось дней']
            number_kvk = df_current_row['Количество карт контроля']
            if type(deadline) != str:
                QMessageBox.about(self, "Рассылка начальнику отдела", f'В выбранном отделе: "{department}" отсутствуют несданные карты контроля')
                return -2
            send_message_heads(name, email, department, deadline, left_days, number_kvk)
            QMessageBox.about(self, "Рассылка начальнику отдела", f'Рассылка отправлена начальнику отдела:\n"{department}".')
            log_print('Сообщения начальнику прошли')
            self.df_general_information.loc[current_row, 'Дата последнего оповещения начальникам отдела'] = dt.datetime.now().strftime('%d.%m.%Y')
            self.fill_table_general_information()
            return 0
        except Exception as e:
            log_print('Сообщения начальнику упали', e)

    def message_kurator(self):
        try:
            kurator_message(self.df_general_information)
            kurators = self.df_general_information.loc[self.df_general_information['Осталось дней'].notna(), 'Куратор'].drop_duplicates().tolist()
            self.df_general_information.loc[self.df_general_information['Куратор'].isin(kurators), 'Дата последнего оповещения'] = dt.datetime.now().strftime('%d.%m.%Y')
            self.fill_table_general_information()
        except Exception as e:
            log_print('Сообщения упали', e)

class KVK_list_tab(QWidget):
    """Вкладка для всех КВК"""
    def __init__(self):

        super().__init__()
        self.departments = pd.read_excel('kadrifile.xlsx', sheet_name='Штатка', skiprows=0, usecols='B')
        self.departments = self.departments.drop_duplicates()['Подразделение'].tolist()

        self.card_inner_control = pd.read_excel('df2_base.xlsx', sheet_name='Карты внутреннего контроля', skiprows=0, usecols='A:F')

        self.card_inner_control.loc[self.card_inner_control['Статус'].isna(), 'Осталось дней'] = \
            self.card_inner_control.loc[self.card_inner_control['Статус'].isna(), 'Срок сдачи']\
                .map(lambda x: (dt.datetime.today() - dt.datetime.strptime(x, '%d.%m.%Y')).days * -1)
        self.card_inner_control['Осталось дней'].astype('Int64')

        self.columnames = list(self.card_inner_control.columns)
        self.target_card_inner_control = self.card_inner_control

        self._GUI_init()

    def _GUI_init(self):
        self.filter_checkbox = QCheckBox()
        self.main_field = QVBoxLayout()
        self.above_condition_field = QHBoxLayout()

        self.field_condition_label = QLabel("Департамент")

        self.field_condition_edit = QComboBox()
        self.field_condition_edit.setEditable(True)
        self.field_condition_edit.addItems(self.departments)
        self.field_condition_edit.currentTextChanged.connect(self.on_department_change)

        # Выбор отдела
        self.above_condition_field.addWidget(self.field_condition_label)
        self.above_condition_field.addWidget(self.field_condition_edit)

        #Нижняя таблица
        self.KVK_table_widget = QTableWidget()
        self.KVK_table_widget.setColumnCount(self.card_inner_control.shape[1])
        self.KVK_table_widget.setHorizontalHeaderLabels(self.card_inner_control.columns.to_list())
        #self.KVK_table_widget.cellChanged.connect(self.handle_cell_changed)

        self._fill_table_widget()

        # Добавление на лэйаут
        self.main_field.addLayout(self.above_condition_field)
        self.main_field.addWidget(self.KVK_table_widget)


        self.setLayout(self.main_field)
        self.show()

    def _pull_from_table_widget(self):
        self.card_inner_control = pd.DataFrame(columns=self.columnames)
        for row in range(self.KVK_table_widget.rowCount()):
            rowData = []
            for col in range(self.KVK_table_widget.columnCount()):
                elem = self.KVK_table_widget.item(row, col).text()
                rowData.append(elem)
        self.card_inner_control = pd.DataFrame([rowData], columns=self.columnames)

    def handle_date_changed(self, date):
        """Хэндлер когда изменяется календарик"""
        row, col = self.KVK_table_widget.currentRow(), self.KVK_table_widget.currentColumn()
        date = date.toPyDate()
        if col == 1:
        #Подсчет оставшихся дней и дедлайн

            deadline = date + dt.timedelta(days=14)
            days_left = (deadline - dt.datetime.now().date()).days
            deadline = deadline.strftime('%d.%m.%Y')

            #Изменение в пандасе
            self.target_card_inner_control.loc[row, 'Дата изменения'] = date.strftime('%d.%m.%Y')
            self.target_card_inner_control.loc[row, 'Осталось дней'] = days_left
            self.target_card_inner_control.loc[row, 'Срок сдачи'] = deadline

            #Изменение в виджете
            self.KVK_table_widget.setItem(row, col + 1, QTableWidgetItem(deadline))
            self.KVK_table_widget.setItem(row, col + 2, QTableWidgetItem(str(days_left)))
        elif col == 4:
            days_left = 0
            status = 'Cдан'

            # изменения пандас
            self.target_card_inner_control.loc[row, 'Осталось дней'] = days_left
            self.target_card_inner_control.loc[row, 'Статус'] = status
            self.target_card_inner_control.loc[row, 'Дата получения карты'] = date.strftime('%d.%m.%Y')
            self.target_card_inner_control.sort_values(by='Статус', inplace=True)


            # ищменения виджета
            self.KVK_table_widget.setItem(row, col - 1, QTableWidgetItem(str(days_left)))
            self.KVK_table_widget.setItem(row, col + 1, QTableWidgetItem(status))
            self.KVK_table_widget.sortItems(5, QtCore.Qt.DescendingOrder)


        #Общий код
        self.KVK_table_widget.resizeColumnsToContents()
        self._merge_dataframes()

    def _fill_table_widget(self):
        self.KVK_table_widget.blockSignals(True)

        self.target_card_inner_control = self.card_inner_control[self.card_inner_control['Отдел'] == self.field_condition_edit.currentText()]
        self.target_card_inner_control.reset_index(drop=True, inplace=True)
        self.KVK_table_widget.setRowCount(0)
        for i, row in self.target_card_inner_control.iterrows():
            self.KVK_table_widget.setRowCount(self.KVK_table_widget.rowCount() + 1)

            for icol in range(self.KVK_table_widget.columnCount()):
                tmp = str(row[icol])
                if icol == 1 or icol == 4:
                    if type(row[icol]) == float:
                        value = dt.datetime.now().strftime('%d.%m.%Y')
                    else:
                        value = row[icol]

                    calendarr = QCalendarWidget()
                    calendarr.setGridVisible(True)
                    calendarr.setStyleSheet("QCalendarWidget QToolButton"
                                            "{"
                                            "background-color : lightgrey;"
                                            "color : black"
                                            "}")

                    calendarchik = QDateEdit(calendarPopup=True)
                    calendarchik.setCalendarWidget(calendarr)
                    calendarchik.setDisplayFormat('dd.MM.yyyy')
                    calendarchik.setDateTime(dt.datetime.strptime(value, '%d.%m.%Y'))
                    calendarchik.setStyleSheet("""
                                    QDateEdit
                                            {
                                            border : none;
                                            }
                    """)
                    calendarchik.dateChanged.connect(self.handle_date_changed)

                    self.KVK_table_widget.setCellWidget(i, icol, calendarchik)

                if tmp != 'nan' and tmp != '<NA>' and tmp != np.nan:
                    self.KVK_table_widget.setItem(i, icol, QTableWidgetItem(tmp))
                else:
                    self.KVK_table_widget.setItem(i, icol, QTableWidgetItem(''))
        self.KVK_table_widget.resizeColumnsToContents()
        self.KVK_table_widget.blockSignals(False)


    def on_department_change(self, value):
        self.target_card_inner_control = self.card_inner_control[self.card_inner_control['Отдел'] == value]
        self._fill_table_widget()

    def _merge_dataframes(self):
        """Соединяет таргет датафрейм с главным"""
        self.card_inner_control.drop(
            self.card_inner_control[self.card_inner_control['Отдел'] == self.field_condition_edit.currentText()].index, inplace=True)
        self.card_inner_control = pd.concat([self.card_inner_control, self.target_card_inner_control], ignore_index=True)

    def add_new_card(self):
        dep = self.field_condition_edit.currentText()
        deadline = dt.datetime.now() + dt.timedelta(days=14)
        deadline = deadline.strftime('%d.%m.%Y')

        self.card_inner_control.loc[len(self.card_inner_control)] = [dep, dt.datetime.now().strftime('%d.%m.%Y'), deadline , 14, np.nan, np.nan]
        self._fill_table_widget()
    def del_last_card(self):
        dep = self.field_condition_edit.currentText()
        self.target_card_inner_control
        self.target_card_inner_control.drop(self.target_card_inner_control.tail(1).index, inplace=True)
        self._merge_dataframes()
        self._fill_table_widget()


stylesheet12 = """
    App1 {
        background-image: url(fon2.png); 
        background-repeat: no-repeat; 
        background-position: center;
    }

    MyWindow {
        background-color: #D7DCEE;
        background-image: url(1234.png);
    }


"""

if __name__ == "__main__":
    if '-t' in sys.argv:
        print('Test open')
        app = QApplication(sys.argv)
        app.setStyleSheet(stylesheet12)
        window = MyWindow()
        sys.exit(app.exec_())
    else:
        app = QApplication([])
        app.setStyleSheet(stylesheet12)
        window = App1()
        app.exec()
