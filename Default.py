import json
import os
import pathlib
import default_window

from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QLineEdit, QDialog, QButtonGroup, QLabel, QSizePolicy, QPushButton


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
        self.name_eng = ['path_old', 'path_new', 'path_file_num',
                         'classified', 'num_scroll', 'list_item', 'number', 'executor', 'conclusion', 'prescription',
                         'print_people', 'date', 'executor_acc_sheet', 'act', 'statement',
                         'account_post', 'account_signature', 'account_path',
                         'firm', 'path_form_27_create',
                         'path_old_print', 'account_numbers', 'path_form_27', 'add_account_num',
                         'HDD_number']
        self.name_rus = ['Путь к исходным файлам', 'Путь к конечным файлам', 'Путь к файлу номеров',
                         'Гриф секретности', 'Номер экземпляра', 'Пункт перечня', 'Номер', 'Протокол', 'Заключение',
                         'Предписание', 'Печать', 'Дата', 'Сопровод', 'Акт', 'Утверждение',
                         'Должность', 'ФИО подпись', 'Путь к описи',
                         'Организация', 'Форма 27 (сохранение)',
                         'Путь к файлам для печати', 'Путь к учетным номерам', 'Форма 27 для печати',
                         'Путь к доп. файлу уч. ном.', 'Номер НЖМД']
        with open(pathlib.Path(self.path_for_default, 'Настройки.txt'), "r", encoding='utf-8-sig') as f:  # Открываем
            self.data = json.load(f)  # Загружаем данные
        self.buttongroup_add = QButtonGroup()
        self.buttongroup_add.buttonClicked[int].connect(self.add_button_clicked)
        self.pushButton_ok.clicked.connect(self.accept)  # Принять
        self.pushButton_cancel.clicked.connect(lambda: self.close())  # Отмена
        self.line = {}  # Для имен
        self.name = {}  # Для значений
        self.button = {}  # Для кнопки «изменить»
        for i, el in enumerate(self.name_rus):  # Заполняем
            self.line[i] = QLabel(self.groupBox_sett)  # Помещаем в фрейм
            self.line[i].setText(el)  # Название элемента
            self.line[i].setFont(QFont("Times", 12, QFont.Light))  # Шрифт, размер
            self.line[i].setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)  # Размеры виджета
            self.line[i].setDisabled(True)  # Делаем неактивным, чтобы нельзя было просто так редактировать
            self.gridLayout_sett.addWidget(self.line[i], i, 0)  # Добавляем виджет
            self.name[i] = Button(self.groupBox_sett)  # Помещаем в фрейм
            try:  # Проверяем есть ли значение
                self.name[i].setText(self.data[self.name_eng[self.name_rus.index(el)]])
            except KeyError:
                pass
            self.name[i].setFont(QFont("Times", 12, QFont.Light))  # Шрифт, размер
            self.name[i].setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)  # Размеры виджета
            self.name[i].setDisabled(True)  # Неактивный
            self.gridLayout_sett.addWidget(self.name[i], i, 1)  # Помещаем в фрейм
            self.button[i] = QPushButton("Изменить", self.groupBox_sett)  # Создаем кнопку
            self.button[i].setFont(QFont("Times", 12, QFont.Light))  # Размер шрифта
            self.button[i].setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)  # Размеры виджета
            self.buttongroup_add.addButton(self.button[i], i)  # Добавляем в группу
            self.gridLayout_sett.addWidget(self.button[i], i, 2)  # Добавляем в фрейм по месту

    def add_button_clicked(self, number):  # Если кликнули по кнопке
        self.name[number].setEnabled(True)  # Делаем активным для изменения

    def accept(self):  # Если нажали кнопку принять
        for el in self.name:  # Пробегаем значения
            if self.name[el].isEnabled():  # Если виджет активный (означает потенциальное изменение)
                if self.name[el].text():  # Если внутри виджета есть текст, то помещаем внутрь базы
                    self.data[self.name_eng[self.name_rus.index(self.line[el].text())]] = self.name[el].text()
                else:  # Если нет текста, то удаляем значение
                    self.data[self.name_eng[self.name_rus.index(self.line[el].text())]] = None
                    # self.data.pop(self.name_eng[self.name_rus.index(self.line[el].text())], None)
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

