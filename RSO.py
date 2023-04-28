import datetime
import os
import sys
import json
import pathlib
import logging

import Main
import about
from Default import DefaultWindow
from AccountNum import AccountNumWindow
from NumberInstance import NumberInstance
from Check import doc_format, doc_print
from PrintDoc import PrintDoc
from FormatDoc import FormatDoc

from PyQt5 import QtPrintSupport

from PyQt5.QtCore import (QDir, QTranslator, QLocale, QLibraryInfo)
from PyQt5.QtWidgets import (QMainWindow, QApplication, QFileDialog, QMessageBox, QDialog, QDesktopWidget)

from queue import Queue


class AboutWindow(QDialog, about.Ui_Dialog):  # Для отображения информации
    def __init__(self):
        super().__init__()
        self.setupUi(self)


def about():  # Открываем окно с описанием
    window_add = AboutWindow()
    window_add.exec_()


def account_number():  # Запускаем окно для создания файла учетных номеров.
    window_add = AccountNumWindow()
    window_add.exec_()


def create_instance():  # Запускаем окно для создания экземпляров.
    window_add = NumberInstance()
    window_add.exec_()


class MainWindow(QMainWindow, Main.Ui_MainWindow):  # Главное окно

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setupUi(self)
        filename = str(datetime.date.today()) + '_logs.log'
        os.makedirs(pathlib.Path('logs'), exist_ok=True)
        filemode = 'a' if pathlib.Path('logs', filename).is_file() else 'w'
        logging.basicConfig(filename=pathlib.Path('logs', filename),
                            level=logging.DEBUG,
                            filemode=filemode,
                            format="%(asctime)s - %(levelname)s - %(funcName)s: %(lineno)d - %(message)s")
        self.queue = Queue()
        self.finish = False  # Для запуска следующего потока в очереди
        self.pushButton_print.clicked.connect(self.printing)  # Кнопка распечатать
        self.pushButton_insert.clicked.connect(self.insert_head_foot)  # Кнопка вставить
        # Баттоны для кнопок открыть
        self.pushButton_folder_old_doc.clicked.connect(lambda: self.browse(self.lineEdit_path_folder_old_doc))
        self.pushButton_folder_new_doc.clicked.connect(lambda: self.browse(self.lineEdit_path_folder_new_doc))
        self.pushButton_file_file_num.clicked.connect(lambda: self.browse(self.lineEdit_path_file_file_num))
        self.pushButton_folder_account.clicked.connect(lambda: self.browse(self.lineEdit_path_folder_account))
        self.pushButton_folder_open_form_27.clicked.connect(lambda:
                                                            self.browse(self.lineEdit_path_folder_form_27_create))
        self.pushButton_folder_old_print.clicked.connect(lambda: self.browse(self.lineEdit_path_folder_old_print))
        self.pushButton_file_form27_print.clicked.connect(lambda: self.browse(self.lineEdit_path_file_form_27_print))
        self.pushButton_file_account_numbers.clicked.connect(lambda:
                                                             self.browse(self.lineEdit_path_file_account_numbers))
        self.pushButton_file_add_account_numbers.clicked.connect(lambda:
                                                                 self.browse(
                                                                     self.lineEdit_path_file_add_account_numbers))
        self.pushButton_folder_sp.clicked.connect(lambda: self.browse(self.lineEdit_path_folder_sp))
        # Для выбора принтера по умолчанию
        self.comboBox_printer.addItems(QtPrintSupport.QPrinterInfo.availablePrinterNames())
        self.comboBox_printer.currentTextChanged.connect(self.text_changed)
        self.lineEdit_printer.setText(QtPrintSupport.QPrinterInfo.defaultPrinterName())
        # Кнопки в меню
        self.action_default.triggered.connect(self.default_settings)
        self.action_instance.triggered.connect(create_instance)
        self.action_about.triggered.connect(about)
        self.action_account_number.triggered.connect(account_number)
        # Группа для кнопок принтера
        self.button_gr = [self.radioButton_last_duplex, self.radioButton_duplex, self.radioButton_one_side]
        # Если изменяем начальный номер
        self.path_for_default = pathlib.Path.cwd()  # Путь для файла настроек
        # Имена в файле
        self.list = {'insert-path_old': ['Путь к исходным файлам', self.lineEdit_path_folder_old_doc],
                     'insert-path_new': ['Путь к конечным файлам', self.lineEdit_path_folder_new_doc],
                     'insert-path_file_num': ['Путь к файлу номеров', self.lineEdit_path_file_file_num],
                     'insert-path_sp': ['Путь к материалам СП', self.lineEdit_path_folder_sp],
                     'data-classified': ['Гриф секретности', self.comboBox_classified],
                     'data-num_scroll': ['Номер экземпляра', self.lineEdit_num_scroll],
                     'data-list_item': ['Пункт перечня', self.lineEdit_list_item],
                     'data-number': ['Номер', self.lineEdit_number],
                     'data-protocol': ['Протокол', self.lineEdit_protocol],
                     'data-conclusion': ['Заключение', self.lineEdit_conclusion],
                     'data-prescription': ['Предписание', self.lineEdit_prescription],
                     'data-print_people': ['Печать', self.lineEdit_print],
                     'data-date': ['Дата', self.lineEdit_date],
                     'data-executor_acc_sheet': ['Сопровод', self.lineEdit_executor_acc_sheet],
                     'data-act': ['Акт', self.lineEdit_act],
                     'data-statement': ['Утверждение', self.lineEdit_statement],
                     'account-account_post': ['Должность', self.lineEdit_account_post],
                     'account-account_signature': ['ФИО подпись', self.lineEdit_account_signature],
                     'account-account_path': ['Путь к описи', self.lineEdit_path_folder_account],
                     'form27-firm': ['Организация', self.lineEdit_firm],
                     'form27-path_form_27_create': ['Форма 27 (вставка)', self.lineEdit_path_folder_form_27_create],
                     'print-path_old_print': ['Путь к файлам для печати', self.lineEdit_path_folder_old_print],
                     'print-account_numbers': ['Путь к учетным номерам', self.lineEdit_path_file_account_numbers],
                     'print-path_form_27': ['Форма 27 (печать)', self.lineEdit_path_file_form_27_print],
                     'print-add_account_num': ['Путь к доп. файлу уч. ном.',
                                               self.lineEdit_path_file_add_account_numbers],
                     'data-HDD_number': 'Номер НЖМД'}
        # Грузим значения по умолчанию
        try:
            with open(pathlib.Path(pathlib.Path.cwd(), 'Настройки.txt'), "r", encoding='utf-8-sig') as f:
                data = json.load(f)
        except FileNotFoundError:
            with open(pathlib.Path(pathlib.Path.cwd(), 'Настройки.txt'), "w", encoding='utf-8-sig') as f:
                json.dump({}, f, ensure_ascii=False, sort_keys=True, indent=4)
                data = {}
        self.hdd_number = None
        self.default_date(data)
        qt_rectangle = self.frameGeometry()
        center_point = QDesktopWidget().availableGeometry().center()
        qt_rectangle.moveCenter(center_point)
        self.move(qt_rectangle.topLeft())
        self.thread = None

    def default_date(self, d):
        for el in self.list:
            if el in d:
                if el == 'data-classified':  # Если элемент гриф секретности
                    index = 0
                    if d[el] is None:
                        self.comboBox_classified.setCurrentIndex(0)
                        continue
                    elif d[el] == 'ДСП':
                        index = 1
                    text_element = ['CC', 'СС', 'C', 'С', 'OB', 'ОВ']  # Названия, которые могут быть (англ. и рус.)
                    for element in text_element:  # Для элементов в списке
                        if d[el] == element:  # Если элемент совпадает, то смотрим что бы он был нечетным
                            if (text_element.index(element) - 1) / 2 < 0 or (text_element.index(element) - 1) % 2 != 0:
                                text = text_element.index(element) + 1  # Выбираем следующий
                                index = self.comboBox_classified.findText(text_element[text])  # Запоминаем индекс
                            else:
                                text = text_element.index(element)
                                index = self.comboBox_classified.findText(text_element[text])
                            break  # Прерываем цикл
                    self.comboBox_classified.setCurrentIndex(index)  # Помещаем соответствующий элемент
                elif el == 'data-HDD_number':
                    self.hdd_number = d[el]
                else:  # Если любой другой элемент
                    self.list[el][1].setText(d[el])  # Помещаем значение

    def default_settings(self):  # Запускаем окно с настройками по умолчанию.
        self.close()
        window_add = DefaultWindow(self, self.path_for_default)
        window_add.show()

    def on_message_changed(self, title, description):  # Для вывода сообщений
        if title == 'УПС!':  # Ошибка
            QMessageBox.critical(self, title, description)
        elif title == 'ВНИМАНИЕ!':  # Предупреждение
            QMessageBox.warning(self, title, description)
        elif title == 'Вопрос':
            ans = QMessageBox.question(self, title, description,
                                       QMessageBox.Cancel | QMessageBox.Ignore | QMessageBox.Retry, QMessageBox.Retry)
            if ans == QMessageBox.Retry:
                self.thread.q.put(2)
            elif ans == QMessageBox.Ignore:
                self.thread.q.put(3)
            else:
                self.thread.q.put(4)

    def browse(self, line_edit):  # Для кнопки открыть
        if 'folder' in self.sender().objectName():
            directory = QFileDialog.getExistingDirectory(self, "Открыть папку", QDir.currentPath())
        else:
            directory = QFileDialog.getOpenFileName(self, "Открыть файл", QDir.currentPath())
        if directory and isinstance(directory, tuple):
            if directory[0]:
                line_edit.setText(directory[0])
        elif directory and isinstance(directory, str):
            line_edit.setText(directory)

    def text_changed(self):  # Если изменился выбор принтера
        self.lineEdit_printer.setText(self.comboBox_printer.currentText())

    def insert_head_foot(self):
        # Проверка введенных данных перед запуском потока
        output = doc_format(self.lineEdit_path_folder_old_doc, self.lineEdit_path_folder_new_doc,
                            self.lineEdit_path_file_file_num,
                            self.radioButton_FSB_df, self.radioButton_FSTEK_df,
                            self.comboBox_classified, self.lineEdit_num_scroll,
                            self.lineEdit_list_item, self.lineEdit_number, self.lineEdit_protocol,
                            self.lineEdit_conclusion, self.lineEdit_prescription, self.lineEdit_print,
                            self.lineEdit_executor_acc_sheet, self.label_protocol, self.label_conclusion,
                            self.label_prescription, self.label_print, self.label_executor_acc_sheet,
                            self.lineEdit_date, self.lineEdit_act, self.lineEdit_statement,
                            self.groupBox_inventory_insert, self.radioButton_40_num,
                            self.radioButton_all_doc, self.lineEdit_account_post,
                            self.lineEdit_account_signature, self.lineEdit_path_folder_account, self.hdd_number,
                            self.groupBox_form27_insert, self.lineEdit_firm, self.lineEdit_path_folder_form_27_create,
                            self.groupBox_instance, self.lineEdit_number_instance, self.checkBox_conclusion_instance,
                            self.checkBox_protocol_instance, self.checkBox_preciption_instance, self.action_package,
                            self.action_report_MO, self.checkBox_folder_path_sp, self.lineEdit_path_folder_sp)
        if type(output) == list:
            self.on_message_changed(output[0], output[1])
            return
        # Если всё прошло запускаем поток
        output['queue'], output['logging'] = self.queue, logging
        self.thread = FormatDoc(output)
        self.thread.progress.connect(self.progressBar.setValue)
        self.thread.status.connect(self.show_mess)
        self.thread.messageChanged.connect(self.on_message_changed)
        self.thread.finished.connect(self.stop_thread)
        self.thread.start()

    def printing(self):
        # Проверка введенных данных перед запуском потока
        output = doc_print(self.radioButton_FSB_print, self.radioButton_FSTEK_print, self.checkBox_conclusion_print,
                           self.checkBox_protocol_print, self.checkBox_preciption_print,
                           self.lineEdit_path_folder_old_print,
                           self.lineEdit_path_file_account_numbers, self.checkBox_file_add_account_numbers,
                           self.lineEdit_path_file_add_account_numbers, self.checkBox_file_form_27,
                           self.lineEdit_path_file_form_27_print,
                           self.button_gr, self.lineEdit_printer, self.checkBox_print_order, self.path_for_default,
                           self.action_package)
        if type(output) == list:
            self.on_message_changed(output[0], output[1])
            return
        # Если всё прошло запускаем поток
        output['logging'] = logging
        self.thread = PrintDoc(output)
        self.thread.progress.connect(self.progressBar.setValue)
        self.thread.status.connect(self.show_mess)
        self.thread.messageChanged.connect(self.on_message_changed)
        self.thread.start()
        self.thread.finished.connect(self.stop_thread)

    def stop_thread(self):  # Завершение потока
        os.chdir(self.path_for_default)

    def show_mess(self, value):  # Вывод значения в статус бар
        self.statusBar().showMessage(value)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    translator = QTranslator(app)
    locale = QLocale.system().name()
    path = QLibraryInfo.location(QLibraryInfo.TranslationsPath)
    translator.load('qtbase_%s' % locale.partition('_')[0], path)
    app.installTranslator(translator)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
