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


def create_instance(self):  # Запускаем окно для создания экземпляров.
    window_add = NumberInstance()
    window_add.exec_()


class MainWindow(QMainWindow, Main.Ui_MainWindow):  # Главное окно

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setupUi(self)
        logging.basicConfig(filename="my_log.log",
                            level=logging.DEBUG,
                            filemode="w",
                            format="%(asctime)s - %(levelname)s - %(funcName)s: %(lineno)d - %(message)s")
        self.q = Queue()
        self.finish = False  # Для запуска следующего потока в очереди
        self.pushButton_print.clicked.connect(self.printing)  # Кнопка распечатать
        self.pushButton_insert.clicked.connect(self.insert_head_foot)  # Кнопка вставить
        # Баттоны для кнопок открыть
        self.pushButton_old.clicked.connect((lambda: self.browse(0)))
        self.pushButton_new.clicked.connect((lambda: self.browse(1)))
        self.pushButton_file_num.clicked.connect((lambda: self.browse(2)))
        self.pushButton_account.clicked.connect((lambda: self.browse(3)))
        self.pushButton_open_form_27.clicked.connect((lambda: self.browse(4)))
        self.pushButton_old_print.clicked.connect((lambda: self.browse(5)))
        self.pushButton_path_form27_print.clicked.connect((lambda: self.browse(6)))
        self.pushButton_account_numbers.clicked.connect((lambda: self.browse(7)))
        self.pushButton_add_account_numbers.clicked.connect((lambda: self.browse(8)))
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
        self.name_eng = ['path_old', 'path_new', 'path_file_num',
                         'classified', 'num_scroll', 'list_item', 'number', 'executor', 'conclusion', 'prescription',
                         'print_people', 'date', 'executor_acc_sheet', 'act', 'statement',
                         'account_post', 'account_signature', 'account_path',
                         'firm', 'path_form_27_create',
                         'path_old_print', 'account_numbers', 'path_form_27', 'add_account_num',
                         'HDD_number']
        # Грузим значения по умолчанию
        try:
            with open(pathlib.Path(pathlib.Path.cwd(), 'Настройки.txt'), "r", encoding='utf-8-sig') as f:
                data = json.load(f)
        except FileNotFoundError:
            with open(pathlib.Path(pathlib.Path.cwd(), 'Настройки.txt'), "w", encoding='utf-8-sig') as f:
                json.dump({}, f, ensure_ascii=False, sort_keys=True, indent=4)
                data = {}

        # Линии для заполнения
        self.line = [self.lineEdit_path_old, self.lineEdit_path_new, self.lineEdit_path_file_num,
                     self.comboBox_classified, self.lineEdit_num_scroll, self.lineEdit_list_item, self.lineEdit_number,
                     self.lineEdit_executor, self.lineEdit_conclusion, self.lineEdit_prescription, self.lineEdit_print,
                     self.lineEdit_date, self.lineEdit_executor_acc_sheet, self.lineEdit_act, self.lineEdit_statement,
                     self.lineEdit_account_post, self.lineEdit_account_signature, self.lineEdit_path_account,
                     self.lineEdit_firm, self.lineEdit_path_form_27_create,
                     self.lineEdit_path_old_print, self.lineEdit_path_account_numbers, self.lineEdit_path_form_27_print,
                     self.lineEdit_path_add_account_numbers]
        self.hdd_number = None
        self.default_date(data)
        qt_rectangle = self.frameGeometry()
        center_point = QDesktopWidget().availableGeometry().center()
        qt_rectangle.moveCenter(center_point)
        self.move(qt_rectangle.topLeft())
        self.thread = None

    def default_date(self, d):
        for el in d:  # Проверяем все загруженные данные
            if el in self.name_eng:
                if el == 'classified':  # Если элемент гриф секретности
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
                elif el == 'HDD_number':
                    self.hdd_number = d[el]
                else:  # Если любой другой элемент
                    self.line[self.name_eng.index(el)].setText(d[el])  # Помещаем значение

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

    def browse(self, num):  # Для кнопки открыть
        directory = None
        if num in [2, 6, 7, 8]:  # Если необходимо открыть файл
            directory = QFileDialog.getOpenFileName(self, "Find Files", QDir.currentPath())
        elif num in [0, 1, 3, 4, 5]:  # Если необходимо открыть директорию
            directory = QFileDialog.getExistingDirectory(self, "Find Files", QDir.currentPath())
        # Список линий
        line = [self.lineEdit_path_old, self.lineEdit_path_new, self.lineEdit_path_file_num,
                self.lineEdit_path_account, self.lineEdit_path_form_27_create,
                self.lineEdit_path_old_print, self.lineEdit_path_form_27_print, self.lineEdit_path_account_numbers,
                self.lineEdit_path_add_account_numbers]
        if directory:  # Если нажать кнопку отркыть в диалоге выбора
            if num in [2, 6, 7, 8]:  # Если файлы
                if directory[0]:  # Если есть файл, чтобы не очищалось поле
                    line[num].setText(directory[0])
            else:  # Если директории
                line[num].setText(directory)

    def text_changed(self):  # Если изменился выбор принтера
        self.lineEdit_printer.setText(self.comboBox_printer.currentText())

    def insert_head_foot(self):
        # Проверка введенных данных перед запуском потока
        output = doc_format(self.lineEdit_path_old, self.lineEdit_path_new, self.lineEdit_path_file_num,
                            self.radioButton_FSB_df,
                            self.radioButton_FSTEK_df, self.comboBox_classified, self.lineEdit_num_scroll,
                            self.lineEdit_list_item, self.lineEdit_number, self.lineEdit_executor,
                            self.lineEdit_conclusion, self.lineEdit_prescription, self.lineEdit_print,
                            self.lineEdit_executor_acc_sheet, self.label_executor, self.label_conclusion,
                            self.label_prescription, self.label_print, self.label_executor_acc_sheet,
                            self.lineEdit_date, self.lineEdit_act, self.lineEdit_statement,
                            self.groupBox_inventory_insert, self.radioButton_40_num,
                            self.radioButton_all_doc, self.lineEdit_account_post,
                            self.lineEdit_account_signature, self.lineEdit_path_account, self.hdd_number,
                            self.groupBox_form27_insert, self.lineEdit_firm, self.lineEdit_path_form_27_create,
                            self.groupBox_instance, self.lineEdit_number_instance, self.checkBoxd_conclusion,
                            self.checkBox_protocol, self.checkBox_preciption, self.action_package,
                            self.action_report_MO)
        if type(output) == list:
            self.on_message_changed(output[0], output[1])
            return
        else:  # Если всё прошло запускаем поток
            output['q'], output['logging'] = self.q, logging
            self.thread = FormatDoc(output)
            self.thread.progress.connect(self.progressBar.setValue)
            self.thread.status.connect(self.show_mess)
            self.thread.messageChanged.connect(self.on_message_changed)
            self.thread.finished.connect(self.stop_thread)
            self.thread.start()

    def printing(self):
        # Проверка введенных данных перед запуском потока
        output = doc_print(self.radioButton_FSB_print, self.radioButton_FSTEK_print, self.lineEdit_path_old_print,
                           self.lineEdit_path_account_numbers, self.checkBox_add_account_numbers,
                           self.lineEdit_path_add_account_numbers, self.checkBox_form_27,
                           self.lineEdit_path_form_27_print,
                           self.button_gr, self.lineEdit_printer, self.checkBox_print_order, self.path_for_default,
                           self.action_package)
        if type(output) == list:
            self.on_message_changed(output[0], output[1])
            return
        else:  # Если всё прошло запускаем поток
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
