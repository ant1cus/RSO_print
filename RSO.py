import datetime
import os
import sys
import json
import pathlib
import logging
import traceback

import Main
import about
from Default import DefaultWindow
from AccountNum import AccountNumWindow
from NumberInstance import NumberInstance
from SortingFile import SortingFile
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
        self.pushButton_file_sp.clicked.connect(lambda: self.browse(self.lineEdit_path_file_sp))
        # Для выбора принтера по умолчанию
        self.comboBox_printer.addItems(QtPrintSupport.QPrinterInfo.availablePrinterNames())
        self.comboBox_printer.currentTextChanged.connect(self.text_changed)
        self.lineEdit_printer.setText(QtPrintSupport.QPrinterInfo.defaultPrinterName())
        # Кнопки в меню
        self.action_default.triggered.connect(self.default_settings)
        self.action_instance.triggered.connect(create_instance)
        self.action_about.triggered.connect(about)
        self.action_account_number.triggered.connect(account_number)
        self.action_sorting.triggered.connect(self.sorting)
        self.action_instruction.triggered.connect(lambda: self.start_document('documents/Инструкция.docx'))
        self.action_registration.triggered.connect(lambda: self.start_document('documents/Номера для регистрации.xlsx'))
        self.action_sp.triggered.connect(lambda: self.start_document('documents/Номера СП.xlsx'))
        # Группа для кнопок принтера
        self.button_gr = [self.radioButton_group4_last_duplex, self.radioButton_group4_duplex,
                          self.radioButton_group4_one_side]
        # Если изменяем начальный номер
        self.path_for_default = pathlib.Path.cwd()  # Путь для файла настроек
        # Имена в файле
        self.list = {'insert-path_folder_old': ['Путь к исходным файлам', self.lineEdit_path_folder_old_doc],
                     'insert-path_folder_new': ['Путь к конечным файлам', self.lineEdit_path_folder_new_doc],
                     'insert-checkBox_file_num': ['Включить файл номеров', self.checkBox_file_num],
                     'insert-path_file_file_num': ['Путь к файлу номеров', self.lineEdit_path_file_file_num],
                     'data-radioButton_group1': ['Ведомство при рег.', [self.radioButton_group1_FSB_df,
                                                                        self.radioButton_group1_FSTEK_df]],
                     'data-comboBox_classified': ['Гриф секретности', self.comboBox_classified,
                                                  ['', 'ДСП', 'С', 'СС', 'ОВ']],
                     'data-num_scroll': ['Номер экземпляра', self.lineEdit_num_scroll],
                     'data-list_item': ['Пункт перечня', self.lineEdit_list_item],
                     'data-checkBox_add_list_item': ['Включить доп. пункт перечня', self.checkBox_add_list_item],
                     'data-add_list_item': ['Доп. пункт перечня', self.lineEdit_add_list_item],
                     'data-number': ['Секретный №', self.lineEdit_number],
                     'data-protocol': ['Протокол', self.lineEdit_protocol],
                     'data-conclusion': ['Заключение', self.lineEdit_conclusion],
                     'data-prescription': ['Предписание', self.lineEdit_prescription],
                     'data-print_people': ['Печать', self.lineEdit_print],
                     'data-date': ['Дата', self.lineEdit_date],
                     'data-executor_acc_sheet': ['Сопровод', self.lineEdit_executor_acc_sheet],
                     'data-act': ['Акт', self.lineEdit_act],
                     'data-statement': ['Утверждение', self.lineEdit_statement],
                     'data-checkBox_conclusion_number': ['Включить номер заключения', self.checkBox_conclusion_number],
                     'data-conclusion_number': ['Номер заключения', self.lineEdit_conclusion_number],
                     'data-add_conclusion_number_date': ['Доп. дата заключения',
                                                         self.lineEdit_add_conclusion_number_date],
                     'sp-groupBox_sp': ['Включить СП', self.groupBox_sp],
                     'sp-path_folder_sp': ['Путь к материалам СП', self.lineEdit_path_folder_sp],
                     'sp-path_file_sp': ['Путь к файлу с номерами', self.lineEdit_path_file_sp],
                     'sp-checkBox_name_gk': ['Включить имя ГК', self.checkBox_name_gk],
                     'sp-lineEdit_name_gk': ['Имя ГК', self.lineEdit_name_gk],
                     'sp-checkBox_conclusion_sp': ['Проверить заключение', self.checkBox_conclusion_sp],
                     'sp-checkBox_protocol_sp': ['Проверить протокол', self.checkBox_protocol_sp],
                     'sp-checkBox_preciption_sp': ['Проверить предписание', self.checkBox_preciption_sp],
                     'sp-checkBox_infocard_sp': ['Проверить инфокарты', self.checkBox_infocard_sp],
                     'account-groupBox_inventory_insert': ['Включить опись', self.groupBox_inventory_insert],
                     'account-radioButton_group2': ['Выбрать кол-во описей', [self.radioButton_group2_40_num,
                                                                              self.radioButton_group2_all_doc]],
                     'account-account_post': ['Должность', self.lineEdit_account_post],
                     'account-account_signature': ['ФИО подпись', self.lineEdit_account_signature],
                     'account-path_folder_account': ['Путь к описи', self.lineEdit_path_folder_account],
                     'form27-groupBox_form27_insert': ['Включить 27 форму', self.groupBox_form27_insert],
                     'form27-firm': ['Организация', self.lineEdit_firm],
                     'form27-path_folder_form_27_create': ['Путь к форме 27', self.lineEdit_path_folder_form_27_create],
                     'instance-groupBox_instance': ['Включить экземпляры', self.groupBox_instance],
                     'instance-checkBox_conclusion_instance': ['Включить заключения',
                                                               self.checkBox_conclusion_instance],
                     'instance-checkBox_protocol_instance': ['Включить протоколы', self.checkBox_protocol_instance],
                     'instance-checkBox_preciption_instance': ['Включить предписания',
                                                               self.checkBox_preciption_instance],
                     'print-radioButton_group3': ['Ведомство при печати', [self.radioButton_group3_FSB_print,
                                                  self.radioButton_group3_FSTEK_print]],
                     'print-checkBox_conclusion_print': ['Включить заключения', self.checkBox_conclusion_print],
                     'print-checkBox_protocol_print': ['Включить протокол', self.checkBox_protocol_print],
                     'print-checkBox_preciption_print': ['Включить предписание', self.checkBox_preciption_print],
                     'print-path_folder_old_print': ['Путь к файлам для печати', self.lineEdit_path_folder_old_print],
                     'print-path_file_account_numbers': ['Путь к учетным номерам',
                                                         self.lineEdit_path_file_account_numbers],
                     'print-checkBox_file_form_27': ['Включить 27 форму', self.checkBox_file_form_27],
                     'print-path_file_form_27': ['Путь к форме 27', self.lineEdit_path_file_form_27_print],
                     'print-checkBox_file_add_account_numbers': ['Включить доп. номера',
                                                                 self.checkBox_file_add_account_numbers],
                     'print-path_file_add_account_num': ['Путь к доп. файлу уч. ном.',
                                                         self.lineEdit_path_file_add_account_numbers],
                     'print-radioButton_group4': ['Метод печати', [self.radioButton_group4_duplex,
                                                                   self.radioButton_group4_last_duplex,
                                                                   self.radioButton_group4_one_side]],
                     'print-checkBox_print_order': ['Включить печать по порядку', self.checkBox_print_order],
                     'data-HDD_number': ['Номер НЖМД']}
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

    def default_date(self, incoming_data):
        for el in self.list:
            if el in incoming_data:
                if el == 'data-classified':  # Если элемент гриф секретности
                    index = 0
                    if incoming_data[el] is None:
                        self.comboBox_classified.setCurrentIndex(0)
                        continue
                    elif incoming_data[el] == 'ДСП':
                        index = 1
                    text_element = ['CC', 'СС', 'C', 'С', 'OB', 'ОВ']  # Названия, которые могут быть (англ. и рус.)
                    for element in text_element:  # Для элементов в списке
                        if incoming_data[el] == element:  # Если элемент совпадает, то смотрим что бы он был нечетным
                            if (text_element.index(element) - 1) / 2 < 0 or (text_element.index(element) - 1) % 2 != 0:
                                text = text_element.index(element) + 1  # Выбираем следующий
                                index = self.comboBox_classified.findText(text_element[text])  # Запоминаем индекс
                            else:
                                text = text_element.index(element)
                                index = self.comboBox_classified.findText(text_element[text])
                            break  # Прерываем цикл
                    self.comboBox_classified.setCurrentIndex(index)  # Помещаем соответствующий элемент
                elif el == 'data-HDD_number':
                    self.hdd_number = incoming_data[el]
                elif 'checkBox' in el or 'groupBox' in el:
                    self.list[el][1].setChecked(True) if incoming_data[el] \
                        else self.list[el][1].setChecked(False)
                elif 'radioButton' in el:
                    for radio, button in zip(incoming_data[el], self.list[el][1]):
                        if radio:
                            button.setChecked(True)
                        else:
                            button.setAutoExclusive(False)
                            button.setChecked(False)
                        button.setAutoExclusive(True)
                elif 'comboBox' in el:
                    for index, combo in enumerate(incoming_data[el]):
                        if combo:
                            self.list[el][1].setCurrentIndex(index)
                else:  # Если любой другой элемент
                    self.list[el][1].setText(incoming_data[el])  # Помещаем значение

    def default_settings(self):  # Запускаем окно с настройками по умолчанию.
        self.close()
        window_add = DefaultWindow(self, self.path_for_default, self.list)
        window_add.show()

    def start_document(self, document):  # Запускаем окно с настройками по умолчанию.
        os.startfile(pathlib.Path(self.path_for_default, document))

    def sorting(self):  # Запускаем окно для сортировки.
        window_add = SortingFile(self, logging)
        window_add.exec_()

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
        try:
            logging.info('----------------Запускаем insert_head_foot----------------')
            logging.info('Проверка данных')
            output = doc_format(self.lineEdit_path_folder_old_doc, self.lineEdit_path_folder_new_doc,
                                self.lineEdit_path_file_file_num,
                                self.radioButton_group1_FSB_df, self.radioButton_group1_FSTEK_df,
                                self.comboBox_classified, self.lineEdit_num_scroll,
                                self.lineEdit_list_item, self.lineEdit_number, self.checkBox_add_list_item,
                                self.lineEdit_add_list_item, self.lineEdit_protocol,
                                self.lineEdit_conclusion, self.lineEdit_prescription, self.lineEdit_print,
                                self.lineEdit_executor_acc_sheet, self.label_protocol, self.label_conclusion,
                                self.label_prescription, self.label_print, self.label_executor_acc_sheet,
                                self.lineEdit_date, self.lineEdit_act, self.lineEdit_statement,
                                self.checkBox_conclusion_number, self.lineEdit_conclusion_number,
                                self.lineEdit_add_conclusion_number_date,
                                self.groupBox_inventory_insert, self.radioButton_group2_40_num,
                                self.radioButton_group2_all_doc, self.lineEdit_account_post,
                                self.lineEdit_account_signature, self.lineEdit_path_folder_account, self.hdd_number,
                                self.groupBox_form27_insert, self.lineEdit_firm,
                                self.lineEdit_path_folder_form_27_create,
                                self.groupBox_instance, self.lineEdit_number_instance,
                                self.checkBox_conclusion_instance,
                                self.checkBox_protocol_instance, self.checkBox_preciption_instance,
                                self.action_package,
                                self.action_report_MO, self.groupBox_sp, self.lineEdit_path_folder_sp,
                                self.checkBox_name_gk, self.lineEdit_name_gk, self.checkBox_conclusion_sp,
                                self.checkBox_protocol_sp, self.checkBox_preciption_sp, self.checkBox_infocard_sp,
                                self.lineEdit_path_file_sp, self.checkBox_file_num)
            if isinstance(output, list):
                logging.info('Обнаружены ошибки данных')
                self.on_message_changed(output[0], output[1])
                return
            # Если всё прошло запускаем поток
            logging.info('Запуск на выполнение')
            output['queue'], output['logging'] = self.queue, logging
            logging.info('Входные данные:')
            log_data = {file: output[file] if file not in ['firm', 'number', 'list_item'] else 'замена'
                        for file in output}
            logging.info(log_data)
            self.thread = FormatDoc(output)
            self.thread.progress.connect(self.progressBar.setValue)
            self.thread.status.connect(self.show_mess)
            self.thread.messageChanged.connect(self.on_message_changed)
            self.thread.finished.connect(self.stop_thread)
            self.thread.start()
        except BaseException as exception:
            logging.error('Ошибка insert_head_foot')
            logging.error(exception)
            logging.error(traceback.format_exc())
            self.on_message_changed('УПС!', 'Неизвестная ошибка при проверке данных для вставки колонтитулов')

    def printing(self):
        # Проверка введенных данных перед запуском потока
        try:
            logging.info('----------------Запускаем printing----------------')
            logging.info('Проверка данных')
            output = doc_print(self.radioButton_group3_FSB_print, self.radioButton_group3_FSTEK_print,
                               self.checkBox_conclusion_print,
                               self.checkBox_protocol_print, self.checkBox_preciption_print,
                               self.lineEdit_path_folder_old_print,
                               self.lineEdit_path_file_account_numbers, self.checkBox_file_add_account_numbers,
                               self.lineEdit_path_file_add_account_numbers, self.checkBox_file_form_27,
                               self.lineEdit_path_file_form_27_print,
                               self.button_gr, self.lineEdit_printer, self.checkBox_print_order, self.path_for_default,
                               self.action_package)
            if isinstance(output, list):
                logging.info('Обнаружены ошибки данных')
                self.on_message_changed(output[0], output[1])
                return
            # Если всё прошло запускаем поток
            logging.info('Запуск на выполнение')
            output['logging'] = logging
            logging.info('Входные данные:')
            logging.info(output)
            self.thread = PrintDoc(output)
            self.thread.progress.connect(self.progressBar.setValue)
            self.thread.status.connect(self.show_mess)
            self.thread.messageChanged.connect(self.on_message_changed)
            self.thread.start()
            self.thread.finished.connect(self.stop_thread)
        except BaseException as exception:
            logging.error('Ошибка printing')
            logging.error(exception)
            logging.error(traceback.format_exc())
            self.on_message_changed('УПС!', 'Неизвестная ошибка при проверке данных для печати')

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
