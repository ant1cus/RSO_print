import json
import os
import pathlib

from PyQt5.QtCore import QDir

import default_window

from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QLineEdit, QDialog, QButtonGroup, QLabel, QSizePolicy, QPushButton, QComboBox, QFileDialog


class Button(QLineEdit):

    def __init__(self, parent):
        super(Button, self).__init__(parent)

        self.setAcceptDrops(True)

    def dragEnterEvent(self, e):

        if e.mimeData().hasUrls():
            e.accept()
        else:
            super(Button, self).dragEnterEvent(e)

    def dragMoveEvent(self, e):

        super(Button, self).dragMoveEvent(e)

    def dropEvent(self, e):

        if e.mimeData().hasUrls():
            for url in e.mimeData().urls():
                self.setText(os.path.normcase(url.toLocalFile()))
                e.accept()
        else:
            super(Button, self).dropEvent(e)


class DefaultWindow(QDialog, default_window.Ui_Dialog):  # Настройки по умолчанию
    def __init__(self, parent, path):
        super().__init__()
        self.setupUi(self)
        self.parent = parent
        self.path_for_default = path
        # Имена на английском и русском
        self.name_list = {'insert-path_folder_old': 'Путь к исходным файлам',
                          'insert-path_folder_new': 'Путь к конечным файлам',
                          'insert-path_file_file_num': 'Путь к файлу номеров',
                          'insert-checkBox_folder_path_sp': 'Включить путь к СП',
                          'insert-path_folder_sp': 'Путь к материалам СП',
                          'data-classified': 'Гриф секретности',
                          'data-num_scroll': 'Номер экземпляра',
                          'data-list_item': 'Пункт перечня',
                          'data-number': 'Номер',
                          'data-protocol': 'Протокол',
                          'data-conclusion': 'Заключение',
                          'data-prescription': 'Предписание',
                          'data-print_people': 'Печать',
                          'data-date': 'Дата',
                          'data-executor_acc_sheet': 'Сопровод',
                          'data-act': 'Акт',
                          'data-statement': 'Утверждение',
                          'account-groupBox_inventory_insert': 'Включить опись',
                          'account-account_post': 'Должность',
                          'account-account_signature': 'ФИО подпись',
                          'account-path_folder_account': 'Путь к описи',
                          'form27-groupBox_form27_insert': 'Включить 27 форму',
                          'form27-firm': 'Организация',
                          'form27-path_folder_form_27_create': 'Путь к форме',
                          'instance-groupBox_instance': 'Включить экземпляры',
                          'instance-checkBox_conclusion_instance': 'Включить заключения',
                          'instance-checkBox_protocol_instance': 'Включить протоколы',
                          'instance-checkBox_preciption_instance': 'Включить предписания',
                          'print-checkBox_conclusion_print': 'Включить заключения',
                          'print-checkBox_protocol_print': 'Включить протокол',
                          'print-checkBox_preciption_print': 'Включить предписание',
                          'print-path_folder_old_print': 'Путь к файлам для печати',
                          'print-path_file_account_numbers': 'Путь к учетным номерам',
                          'print-checkBox_file_form_27': 'Включить 27 форму',
                          'print-path_file_form_27': 'Путь к форме',
                          'print-checkBox_file_add_account_numbers': 'Включить доп. номера',
                          'print-path_file_add_account_num': 'Путь к доп. файлу уч. ном.',
                          'data-HDD_number': 'Номер НЖМД'}
        self.name_box = [self.groupBox_catalog_insert_default, self.groupBox_data_default,
                         self.groupBox_form_27_default, self.groupBox_inventory_default, self.groupBox_instance,
                         self.groupBox_catalog_print_default]
        self.name_grid = [self.gridLayout_catalog, self.gridLayout_data, self.gridLayout_form_27,
                          self.gridLayout_inventory, self.gridLayout_instance, self.gridLayout_print]
        with open(pathlib.Path(self.path_for_default, 'Настройки.txt'), "r", encoding='utf-8-sig') as f:  # Открываем
            self.data = json.load(f)  # Загружаем данные
        self.buttongroup_add = QButtonGroup()
        self.buttongroup_add.buttonClicked[int].connect(self.add_button_clicked)
        self.buttongroup_clear = QButtonGroup()
        self.buttongroup_clear.buttonClicked[int].connect(self.clear_button_clicked)
        self.buttongroup_open = QButtonGroup()
        self.buttongroup_open.buttonClicked[int].connect(self.open_button_clicked)
        self.pushButton_ok.clicked.connect(self.accept)  # Принять
        self.pushButton_cancel.clicked.connect(lambda: self.close())  # Отмена
        self.line = {}  # Для имен
        self.name = {}  # Для значений
        self.combo = {}  # Для комбобоксов
        self.button = {}  # Для кнопки «изменить»
        self.button_clear = {}  # Для кнопки «очистить»
        self.button_open = {}  # Для кнопки «открыть»
        for i, el in enumerate(self.name_list):  # Заполняем
            frame = grid = False
            for j, n in enumerate(['insert', 'data', 'form27', 'account', 'instance', 'print']):
                if n in el.partition('-')[0]:
                    frame, grid = self.name_box[j], self.name_grid[j]
                    break
            self.line[i] = QLabel(frame)  # Помещаем в фрейм
            self.line[i].setText(self.name_list[el])  # Название элемента
            self.line[i].setFont(QFont("Times", 12, QFont.Light))  # Шрифт, размер
            self.line[i].setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)  # Размеры виджета
            self.line[i].setFixedWidth(200)
            self.line[i].setDisabled(True)  # Делаем неактивным, чтобы нельзя было просто так редактировать
            grid.addWidget(self.line[i], i, 0)  # Добавляем виджет
            if 'checkBox' not in el and 'groupBox' not in el:
                self.button[i] = QPushButton("Изменить", frame)  # Создаем кнопку
                self.button[i].setFont(QFont("Times", 12, QFont.Light))  # Размер шрифта
                self.button[i].setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)  # Размеры виджета
                self.buttongroup_add.addButton(self.button[i], i)  # Добавляем в группу
                grid.addWidget(self.button[i], i, 1)  # Добавляем в фрейм по месту
                self.button_clear[i] = QPushButton("Очистить", frame)  # Создаем кнопку
                self.button_clear[i].setFont(QFont("Times", 12, QFont.Light))  # Размер шрифта
                self.button_clear[i].setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)  # Размеры виджета
                self.buttongroup_clear.addButton(self.button_clear[i], i)  # Добавляем в группу
                grid.addWidget(self.button_clear[i], i, 2)  # Добавляем в фрейм по месту

                self.name[i] = Button(frame)  # Помещаем в фрейм
                if el in self.data:
                    self.name[i].setText(self.data[el])
                self.name[i].setFont(QFont("Times", 12, QFont.Light))  # Шрифт, размер
                self.name[i].setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)  # Размеры виджета
                self.name[i].setStyleSheet("QLineEdit {"
                                           "border-style: solid;"
                                           "}")
                self.name[i].setDisabled(True)  # Неактивный
                grid.addWidget(self.name[i], i, 3)  # Помещаем в фрейм
                if 'Путь' in self.line[i].text():
                    self.button_open[i] = QPushButton("Открыть", frame)  # Создаем кнопку
                    self.button_open[i].setFont(QFont("Times", 12, QFont.Light))  # Размер шрифта
                    self.button_open[i].setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)  # Размеры виджета
                    self.button_open[i].setDisabled(True)  # Неактивный
                    self.buttongroup_open.addButton(self.button_open[i], i)  # Добавляем в группу
                    grid.addWidget(self.button_open[i], i, 4)  # Добавляем в фрейм по месту
            else:
                self.combo[i] = QComboBox(frame)  # Помещаем в фрейм
                self.combo[i].addItems(['Включён', 'Выключен'])
                self.combo[i].setCurrentIndex(0) if el in self.data and self.data[el] \
                    else self.combo[i].setCurrentIndex(1)
                self.combo[i].setFont(QFont("Times", 12, QFont.Light))  # Шрифт, размер
                self.combo[i].setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)  # Размеры виджета
                grid.addWidget(self.combo[i], i, 3)  # Помещаем в фрейм

            # self.name[i] = Button(frame)  # Помещаем в фрейм
            # if el in self.data:
            #     self.name[i].setText(self.data[el])
            # self.name[i].setFont(QFont("Times", 12, QFont.Light))  # Шрифт, размер
            # self.name[i].setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)  # Размеры виджета
            # self.name[i].setDisabled(True)  # Неактивный
            # grid.addWidget(self.name[i], i, 1)  # Помещаем в фрейм
            # self.button[i] = QPushButton("Изменить", frame)  # Создаем кнопку
            # self.button[i].setFont(QFont("Times", 12, QFont.Light))  # Размер шрифта
            # self.button[i].setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)  # Размеры виджета
            # self.buttongroup_add.addButton(self.button[i], i)  # Добавляем в группу
            # grid.addWidget(self.button[i], i, 2)  # Добавляем в фрейм по месту

    def open_button_clicked(self, num):  # Для кнопки открыть
        file = False
        if 'дир.' in self.line[num].text():
            directory = QFileDialog.getExistingDirectory(self, "Открыть папку", QDir.currentPath())
        else:
            directory = QFileDialog.getOpenFileName(self, "Открыть файл", QDir.currentPath())
            file = True
        if directory:  # Если нажать кнопку отркыть в диалоге выбора
            if file:  # Если файлы
                if directory[0]:  # Если есть файл, чтобы не очищалось поле
                    self.name[num].setText(directory[0])
            else:  # Если директории
                self.name[num].setText(directory)

    def add_button_clicked(self, number):  # Если кликнули по кнопке
        self.name[number].setEnabled(True)  # Делаем активным для изменения
        if number in self.button_open:
            self.button_open[number].setEnabled(True)  # Неактивный
        self.name[number].setStyleSheet("QLineEdit {"
                                        "border-style: solid;"
                                        "border-width: 1px;"
                                        "border-color: black; "
                                        "}")

    def clear_button_clicked(self, number):
        self.name[number].clear()

    def accept(self):  # Если нажали кнопку принять
        for i, el in enumerate(self.name_list):  # Пробегаем значения
            if 'checkBox' in el or 'groupBox' in el:
                self.data[el] = True if self.combo[i].currentIndex() == 0 else False
            else:
                if self.name[i].isEnabled():  # Если виджет активный (означает потенциальное изменение)
                    if self.name[i].text():  # Если внутри виджета есть текст, то помещаем внутрь базы
                        self.data[el] = self.name[i].text()
                    else:  # Если нет текста, то удаляем значение
                        self.data[el] = None
        with open(pathlib.Path(self.path_for_default, 'Настройки.txt'), 'w', encoding='utf-8-sig') as f:  # Пишем в файл
            json.dump(self.data, f, ensure_ascii=False, sort_keys=True, indent=4)
        self.close()  # Закрываем

    def closeEvent(self, event):
        os.chdir(pathlib.Path.cwd())
        if self.sender() and self.sender().text() == 'Принять':
            event.accept()
            with open(pathlib.Path(self.path_for_default, 'Настройки.txt'), "r", encoding='utf-8-sig') as f:
                data = json.load(f)  # Загружаем данные
            self.parent.default_date(data)
            self.parent.show()
        else:
            event.accept()
            self.parent.show()

