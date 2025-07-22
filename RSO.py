import os
import queue
import sys
import pathlib
import logging
import Main
import about

from general_function import browse, default_settings, default_data, rewrite_settings, start_thread
from AccountNum import AccountNumWindow
from NumberInstance import NumberInstance
from SortingFile import SortingFile
from Check import doc_format, doc_print
from StartThread import StartThreading
from format_docs import format_doc
from print_docs import print_docs

from PyQt5 import QtPrintSupport

from PyQt5.QtCore import (QTranslator, QLocale, QLibraryInfo)
from PyQt5.QtWidgets import (QMainWindow, QApplication, QDialog)


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
        self.default_path = pathlib.Path.cwd()
        self.mode_description = {'insert_main': {'mode_name': 'insert_main',
                                                 'title': 'Регистрация документов (основная программа) в папке',
                                                 'cancel': 'Регистрация документов (основная программа) в папке'
                                                           ' «name_dir» отменена пользователем',
                                                 'exception': 'Регистрация документов (основная программа) в папке'
                                                              ' «name_dir» не завершена из-за ошибки',
                                                 'success': 'Регистрация документов (основная программа) в папке'
                                                            ' «name_dir» успешно завершена',
                                                 'error': 'Регистрация документов (основная программа) в папке'
                                                          ' «name_dir» завершена с ошибками'
                                                 },
                                 'print_main': {'mode_name': 'print_main',
                                                'title': 'Печать документов (основная программа) в папке',
                                                'cancel': 'Печать документов (основная программа)в папке «name_dir»'
                                                          ' отменена пользователем',
                                                'exception': 'Печать документов (основная программа) в папке «name_dir»'
                                                             ' не завершена из-за ошибки',
                                                'success': 'Печать документов (основная программа) в папке «name_dir»'
                                                           ' успешно завершена',
                                                'error': 'Печать документов (основная программа) в папке «name_dir»'
                                                         ' завершена с ошибками'
                                                },
                                 'insert_41101': {'mode_name': 'insert_41101',
                                                  'title': 'Регистрация документов (программа для 41101) в папке',
                                                  'cancel': 'Регистрация документов (программа для 41101) в папке'
                                                            ' «name_dir» отменена пользователем',
                                                  'exception': 'Регистрация документов (программа для 41101) в папке'
                                                               ' «name_dir» не завершена из-за ошибки',
                                                  'success': 'Регистрация документов (программа для 41101) в папке'
                                                            ' «name_dir» успешно завершена',
                                                  'error': 'Регистрация документов (программа для 41101) в папке'
                                                           ' «name_dir» завершена с ошибками'
                                                  },
                                 'print_41101': {'mode_name': 'print_41101',
                                                 'title': 'Печать документов (программа для 41101) в папке',
                                                 'cancel': 'Печать документов (программа для 41101)в папке «name_dir»'
                                                           ' отменена пользователем',
                                                 'exception': 'Печать документов (программа для 41101) в папке'
                                                              ' «name_dir» не завершена из-за ошибки',
                                                 'success': 'Печать документов (программа для 41101) в папке «name_dir»'
                                                            ' успешно завершена',
                                                 'error': 'Печать документов (программа для 41101) в папке «name_dir»'
                                                          ' завершена с ошибками'
                                                 },
                                 }
        self.pushButton_main_start_path_insert_dir.clicked.connect(lambda:
                                                                   browse(self,
                                                                          self.pushButton_main_start_path_insert_dir,
                                                                          self.lineEdit_main_start_path_insert_dir,
                                                                          self.default_path))
        self.pushButton_main_finish_path_insert_dir.clicked.connect(lambda:
                                                                    browse(self,
                                                                           self.pushButton_main_finish_path_insert_dir,
                                                                           self.lineEdit_main_finish_path_insert_dir,
                                                                           self.default_path))
        self.pushButton_main_file_num.clicked.connect(lambda: browse(self, self.pushButton_main_file_num,
                                                                     self.lineEdit_main_file_num_path,
                                                                     self.default_path))
        self.pushButton_main_path_signature_dir.clicked.connect(lambda: browse(self,
                                                                               self.pushButton_main_path_signature_dir,
                                                                               self.lineEdit_main_path_signature_dir,
                                                                               self.default_path))
        self.pushButton_main_account_path_dir.clicked.connect(lambda: browse(self,
                                                                             self.pushButton_main_account_path_dir,
                                                                             self.lineEdit_main_account_path_dir,
                                                                             self.default_path))
        self.pushButton_main_form27_path_dir.clicked.connect(lambda:
                                                             browse(self, self.pushButton_main_form27_path_dir,
                                                                    self.lineEdit_main_form27_path_dir,
                                                                    self.default_path))
        self.pushButton_main_folder_sp_dir.clicked.connect(lambda: browse(self, self.pushButton_main_folder_sp_dir,
                                                                          self.lineEdit_main_sp_path_dir,
                                                                          self.default_path))
        self.pushButton_main_file_sp.clicked.connect(lambda: browse(self, self.pushButton_main_file_sp,
                                                                    self.lineEdit_main_file_sp_path, self.default_path))
        self.pushButton_main_start_path_print_dir.clicked.connect(lambda:
                                                                  browse(self,
                                                                         self.pushButton_main_start_path_print_dir,
                                                                         self.lineEdit_main_start_path_print_dir,
                                                                         self.default_path))
        self.pushButton_main_file_form27_print.clicked.connect(lambda:
                                                               browse(self, self.pushButton_main_file_form27_print,
                                                                      self.lineEdit_main_path_file_form27_print,
                                                                      self.default_path))
        self.pushButton_main_file_account_numbers.clicked.connect(lambda:
                                                                  browse(self,
                                                                         self.pushButton_main_file_account_numbers,
                                                                         self.lineEdit_main_file_account_numbers_path,
                                                                         self.default_path))
        self.pushButton_main_add_account_numbers.clicked.connect(lambda:
                                                                 browse(self, self.pushButton_main_add_account_numbers,
                                                                        self.lineEdit_main_add_account_numbers_path,
                                                                        self.default_path))
        self.pushButton_41101_start_path_insert_dir.clicked.connect(lambda:
                                                                    browse(self,
                                                                           self.pushButton_41101_start_path_insert_dir,
                                                                           self.lineEdit_41101_start_path_insert_dir,
                                                                           self.default_path))
        self.pushButton_41101_finish_path_insert_dir.clicked.connect(lambda:
                                                                     browse(self,
                                                                            self.pushButton_41101_finish_path_insert_dir,
                                                                            self.lineEdit_41101_finish_path_insert_dir,
                                                                            self.default_path))
        self.pushButton_41101_file_num.clicked.connect(lambda: browse(self, self.pushButton_41101_file_num,
                                                                      self.lineEdit_41101_file_num_path,
                                                                      self.default_path))
        self.pushButton_41101_account_path_dir.clicked.connect(lambda: browse(self,
                                                                              self.pushButton_41101_account_path_dir,
                                                                              self.lineEdit_41101_account_path_dir,
                                                                              self.default_path))
        self.pushButton_41101_form27_path_dir.clicked.connect(lambda:
                                                              browse(self, self.pushButton_41101_form27_path_dir,
                                                                     self.lineEdit_41101_form27_path_dir,
                                                                     self.default_path))
        self.pushButton_41101_start_path_print_dir.clicked.connect(lambda:
                                                                   browse(self,
                                                                          self.pushButton_41101_start_path_print_dir,
                                                                          self.lineEdit_41101_start_path_print_dir,
                                                                          self.default_path))
        self.pushButton_41101_file_form27_print.clicked.connect(lambda:
                                                                browse(self, self.pushButton_41101_file_form27_print,
                                                                       self.lineEdit_41101_path_file_form27_print,
                                                                       self.default_path))
        self.pushButton_41101_file_account_numbers.clicked.connect(lambda:
                                                                   browse(self,
                                                                          self.pushButton_41101_file_account_numbers,
                                                                          self.lineEdit_41101_file_account_numbers_path,
                                                                          self.default_path))
        self.pushButton_41101_add_account_numbers.clicked.connect(lambda:
                                                                  browse(self,
                                                                         self.pushButton_41101_add_account_numbers,
                                                                         self.lineEdit_41101_add_account_numbers_path,
                                                                         self.default_path))
        # Для выбора принтера по умолчанию
        self.comboBox_main_printer.addItems(QtPrintSupport.QPrinterInfo.availablePrinterNames())
        self.comboBox_main_printer.currentTextChanged.connect(self.text_changed)
        self.lineEdit_main_printer.setText(QtPrintSupport.QPrinterInfo.defaultPrinterName())
        # Группа для кнопок принтера
        self.button_gr = [self.radioButton_main_group4_last_duplex, self.radioButton_main_group4_duplex,
                          self.radioButton_main_group4_one_side]
        # Имена в файле
        self.lines = {'insertMain-start_path_insert': ['Путь к исходным файлам',
                                                       self.lineEdit_main_start_path_insert_dir],
                      'insertMain-finish_path_insert': ['Путь к конечным файлам',
                                                        self.lineEdit_main_finish_path_insert_dir],
                      'insertMain-checkBox_file_num': ['Включить файл номеров', self.checkBox_main_file_num],
                      'insertMain-file_num_path': ['Путь к файлу номеров', self.lineEdit_main_file_num_path],
                      'insertMain-checkBox_signature': ['Включить подписи', self.checkBox_main_signature],
                      'insertMain-path_signature_dir': ['Путь подписям', self.lineEdit_main_path_signature_dir],
                      'insertMain-radioButton_group1': ['Ведомство при рег.',
                                                        [self.radioButton_main_group1_FSB_df,
                                                         self.radioButton_main_group1_FSTEK_df]],
                      'insertMain-comboBox_classified': ['Гриф секретности', self.comboBox_main_classified,
                                                         ['', 'ДСП', 'С', 'СС', 'ОВ']],
                      'insertMain-num_scroll': ['Номер экземпляра', self.lineEdit_main_num_scroll],
                      'insertMain-list_item': ['Пункт перечня', self.lineEdit_main_list_item],
                      'insertMain-checkBox_add_list_item': ['Включить доп. пункт перечня',
                                                            self.checkBox_main_add_list_item],
                      'insertMain-add_list_item': ['Доп. пункт перечня', self.lineEdit_main_add_list_item],
                      'insertMain-secret_number': ['Секретный №', self.lineEdit_main_secret_number],
                      'insertMain-HDD_number': ['Номер НЖМД', self.lineEdit_main_HDD_number],
                      'insertMain-telephone': ['Номер телефона', self.lineEdit_main_telephone],
                      'insertMain-conclusion_executor': ['Исп. заключение', self.lineEdit_main_conclusion_executor],
                      'insertMain-checkBox_conclusion_number': ['Включить доп номер заключения',
                                                                self.checkBox_main_conclusion_number],
                      'insertMain-conclusion_number': ['Доп номер заключения', self.lineEdit_main_conclusion_number],
                      'insertMain-dateEdit_add_conclusion_date': ['Доп. дата заключения',
                                                                  self.dateEdit_main_add_conclusion_date],
                      'insertMain-protocol_executor': ['Исп. протокол', self.lineEdit_main_protocol_executor],
                      'insertMain-prescription_executor': ['Исп. предписание',
                                                           self.lineEdit_main_prescription_executor],
                      'insertMain-acc_sheet_executor': ['Исп. сопровод', self.lineEdit_main_acc_sheet_executor],
                      'insertMain-telephone_acc_sheet_executor': ['Тел. исп. сопровода',
                                                                  self.lineEdit_main_telephone_acc_sheet_executor],
                      'insertMain-add_telephone': ['Доб. номер', self.lineEdit_main_add_telephone],
                      'insertMain-print_executor': ['Исп. печать', self.lineEdit_main_print_executor],
                      'insertMain-dateEdit_exec_date': ['Дата', self.dateEdit_main_exec_date],
                      'insertMain-inventory_executor': ['Исп. опись', self.lineEdit_main_inventory_executor],
                      'insertMain-application_executor': ['Исп. приложение', self.lineEdit_main_application_executor],
                      'insertMain-act_executor': ['Исп. акт', self.lineEdit_main_act_executor],
                      'insertMain-statement_executor': ['Исп. утвержд.', self.lineEdit_main_statement_executor],
                      'addInsertMain-groupBox_inventory_insert': ['Включить опись',
                                                                  self.groupBox_main_inventory_insert],
                      'addInsertMain-radioButton_group2': ['Выбрать кол-во описей',
                                                           [self.radioButton_main_group2_40_num,
                                                            self.radioButton_main_group2_all_doc]],
                      'addInsertMain-account_position': ['Должность', self.lineEdit_main_account_position],
                      'addInsertMain-account_executor': ['ФИО опись', self.lineEdit_main_account_executor],
                      'addInsertMain-account_path_folder': ['Путь к описи', self.lineEdit_main_account_path_dir],
                      'addInsertMain-groupBox_form27_insert': ['Включить 27 форму', self.groupBox_main_form27_insert],
                      'addInsertMain-form27_firm': ['Организация', self.lineEdit_main_form27_firm],
                      'addInsertMain-form27_create_path': ['Путь к форме 27', self.lineEdit_main_form27_path_dir],
                      'addInsertMain-groupBox_sp': ['Включить сортировку материалов', self.groupBox_main_sp],
                      'addInsertMain-sp_path_dir': ['Путь к материалам СП', self.lineEdit_main_sp_path_dir],
                      'addInsertMain-file_sp_path': ['Путь к файлу с номерами', self.lineEdit_main_file_sp_path],
                      'addInsertMain-checkBox_name_gk': ['Включить имя ГК', self.checkBox_main_name_gk],
                      'addInsertMain-name_gk': ['Имя ГК', self.lineEdit_main_name_gk],
                      'addInsertMain-checkBox_conclusion_sp': ['Проверить заключение',
                                                               self.checkBox_main_conclusion_sp],
                      'addInsertMain-checkBox_protocol_sp': ['Проверить протокол', self.checkBox_main_protocol_sp],
                      'addInsertMain-checkBox_prescription_sp': ['Проверить предписание',
                                                                 self.checkBox_main_prescription_sp],
                      'addInsertMain-checkBox_infocard_sp': ['Проверить инфокарты', self.checkBox_main_infocard_sp],
                      'addInsertMain-groupBox_instance': ['Включить экземпляры', self.groupBox_main_instance],
                      'addInsertMain-main_number_instance': ['Номера экземпляров', self.lineEdit_main_number_instance],
                      'addInsertMain-checkBox_conclusion_instance': ['Включить заключения',
                                                                     self.checkBox_main_conclusion_instance],
                      'addInsertMain-checkBox_protocol_instance': ['Включить протоколы',
                                                                   self.checkBox_main_protocol_instance],
                      'addInsertMain-checkBox_prescription_instance': ['Включить предписания',
                                                                       self.checkBox_main_prescription_instance],
                      'printMain-radioButton_group3': ['Ведомство при печати',
                                                       [self.radioButton_main_group3_FSB_print,
                                                        self.radioButton_main_group3_FSTEK_print]],
                      'printMain-checkBox_conclusion_print': ['Включить заключения',
                                                              self.checkBox_main_conclusion_print],
                      'printMain-checkBox_protocol_print': ['Включить протокол', self.checkBox_main_protocol_print],
                      'printMain-checkBox_prescription_print': ['Включить предписание',
                                                                self.checkBox_main_prescription_print],
                      'printMain-start_path_print': ['Путь к файлам для печати',
                                                     self.lineEdit_main_start_path_print_dir],
                      'printMain-file_account_numbers_path': ['Путь к учетным номерам',
                                                              self.lineEdit_main_file_account_numbers_path],
                      'printMain-checkBox_add_account_numbers': ['Включить доп. номера',
                                                                 self.checkBox_main_add_account_numbers],
                      'printMain-add_account_numbers_path': ['Путь к доп. файлу уч. ном.',
                                                             self.lineEdit_main_add_account_numbers_path_file],
                      'printMain-checkBox_file_form27': ['Включить 27 форму', self.checkBox_main_file_form27],
                      'printMain-path_file_form27_print': ['Путь к форме 27',
                                                           self.lineEdit_main_path_file_form27_print],
                      'printMain-radioButton_group4': ['Метод печати', [self.radioButton_main_group4_duplex,
                                                                        self.radioButton_main_group4_last_duplex,
                                                                        self.radioButton_main_group4_one_side]],
                      'printMain-checkBox_print_order': ['Включить печать по порядку', self.checkBox_main_print_order],
                      'insert41101-start_path_insert': ['Путь к исходным файлам',
                                                        self.lineEdit_41101_start_path_insert_dir],
                      'insert41101-finish_path_insert': ['Путь к конечным файлам',
                                                         self.lineEdit_41101_finish_path_insert_dir],
                      'insert41101-checkBox_file_num': ['Включить файл номеров', self.checkBox_41101_file_num],
                      'insert41101-file_num_path': ['Путь к файлу номеров', self.lineEdit_41101_file_num_path],
                      'insert41101-radioButton_group5': ['Ведомство при рег.',
                                                         [self.radioButton_41101_group5_FSB_df,
                                                          self.radioButton_41101_group5_FSTEK_df]],
                      'insert41101-comboBox_classified': ['Гриф секретности', self.comboBox_41101_classified,
                                                          ['', 'ДСП', 'С', 'СС', 'ОВ']],
                      'insert41101-num_scroll': ['Номер экземпляра', self.lineEdit_41101_num_scroll],
                      'insert41101-list_item': ['Пункт перечня', self.lineEdit_41101_list_item],
                      'insert41101-checkBox_add_list_item': ['Включить доп. пункт перечня',
                                                             self.checkBox_41101_add_list_item],
                      'insert41101-add_list_item': ['Доп. пункт перечня', self.lineEdit_41101_add_list_item],
                      'insert41101-secret_number': ['Секретный №', self.lineEdit_41101_secret_number],
                      'insert41101-HDD_number': ['Номер НЖМД', self.lineEdit_41101_HDD_number],
                      'insert41101-telephone': ['Номер телефона', self.lineEdit_41101_telephone],
                      'insert41101-conclusion_executor': ['Исп. заключение', self.lineEdit_41101_conclusion_executor],
                      'insert41101-checkBox_conclusion_number': ['Включить доп номер заключения',
                                                                 self.checkBox_41101_conclusion_number],
                      'insert41101-conclusion_number': ['Доп номер заключения', self.lineEdit_41101_conclusion_number],
                      'insert41101-dateEdit_add_conclusion_date': ['Доп. дата заключения',
                                                                   self.dateEdit_41101_add_conclusion_date],
                      'insert41101-protocol_executor': ['Исп. протокол', self.lineEdit_41101_protocol_executor],
                      'insert41101-prescription_executor': ['Исп. предписание',
                                                            self.lineEdit_41101_prescription_executor],
                      'insert41101-acc_sheet_executor': ['Исп. сопровод', self.lineEdit_41101_acc_sheet_executor],
                      'insert41101-telephone_acc_sheet_executor': ['Тел. исп. сопровода',
                                                                   self.lineEdit_main_telephone_acc_sheet_executor],
                      'insert41101-add_telephone': ['Доб. номер', self.lineEdit_41101_add_telephone],
                      'insert41101-print_executor': ['Исп. печать', self.lineEdit_41101_print_executor],
                      'insert41101-dateEdit_exec_date': ['Дата', self.dateEdit_41101_exec_date],
                      'insert41101-inventory_executor': ['Исп. опись', self.lineEdit_41101_inventory_executor],
                      'insert41101-application_executor': ['Исп. приложение', self.lineEdit_41101_application_executor],
                      'insert41101-act_executor': ['Исп. акт', self.lineEdit_41101_act_executor],
                      'insert41101-statement_executor': ['Исп. утвержд.', self.lineEdit_41101_statement_executor],
                      'addInsert41101-groupBox_inventory_insert': ['Включить опись',
                                                                   self.groupBox_41101_inventory_insert],
                      'addInsert41101-radioButton_group6': ['Выбрать кол-во описей',
                                                            [self.radioButton_41101_group6_40_num,
                                                             self.radioButton_41101_group6_all_doc]],
                      'addInsert41101-account_position': ['Должность', self.lineEdit_41101_account_position],
                      'addInsert41101-account_executor': ['ФИО подпись', self.lineEdit_41101_account_executor],
                      'addInsert41101-account_path_folder': ['Путь к описи', self.lineEdit_41101_account_path_dir],
                      'addInsert41101-groupBox_form27_insert': ['Включить 27 форму', self.groupBox_41101_form27_insert],
                      'addInsert41101-form27_firm': ['Организация', self.lineEdit_41101_form27_firm],
                      'addInsert41101-form27_create_path': ['Путь к форме 27', self.lineEdit_41101_form27_path_dir],
                      'print41101-radioButton_group7': ['Ведомство при печати',
                                                        [self.radioButton_41101_group7_FSB_print,
                                                         self.radioButton_41101_group7_FSTEK_print]],
                      'print41101-checkBox_conclusion_print': ['Включить заключения',
                                                               self.checkBox_41101_conclusion_print],
                      'print41101-checkBox_protocol_print': ['Включить протокол', self.checkBox_41101_protocol_print],
                      'print41101-checkBox_prescription_print': ['Включить предписание',
                                                                 self.checkBox_41101_prescription_print],
                      'print41101-start_path_print': ['Путь к файлам для печати',
                                                      self.lineEdit_41101_start_path_print_dir],
                      'print41101-file_account_numbers_path': ['Путь к учетным номерам',
                                                               self.lineEdit_41101_file_account_numbers_path],
                      'print41101-checkBox_add_account_numbers': ['Включить доп. номера',
                                                                  self.checkBox_41101_add_account_numbers],
                      'print41101-add_account_numbers_path': ['Путь к доп. файлу уч. ном.',
                                                              self.lineEdit_41101_add_account_numbers_path_file],
                      'print41101-checkBox_file_form27': ['Включить 27 форму', self.checkBox_41101_file_form27],
                      'print41101-path_file_form27_print': ['Путь к форме 27',
                                                            self.lineEdit_41101_path_file_form27_print],
                      'print41101-radioButton_group8': ['Метод печати', [self.radioButton_41101_group8_duplex,
                                                                         self.radioButton_41101_group8_last_duplex,
                                                                         self.radioButton_41101_group8_one_side]],
                      }
        # Кнопки запуска
        self.pushButton_main_insert.clicked.connect(self.insert_main)
        self.pushButton_41101_insert.clicked.connect(self.insert_41101)
        self.pushButton_main_print.clicked.connect(self.print_main)
        # Кнопки в меню
        self.action_default.triggered.connect((lambda: default_settings(self, self.default_path, self.lines)))
        self.action_instance.triggered.connect(create_instance)
        self.action_about.triggered.connect(about)
        self.action_account_number.triggered.connect(account_number)
        self.action_sorting.triggered.connect(self.sorting)
        self.action_instruction.triggered.connect(lambda: self.start_document('documents/Инструкция.docx'))
        self.action_registration.triggered.connect(lambda: self.start_document('documents/Номера для регистрации.xlsx'))
        self.action_sp.triggered.connect(lambda: self.start_document('documents/Номера СП.xlsx'))
        self.default_data = rewrite_settings(self.default_path)
        self.data = self.default_data["widget_settings"]
        default_data(self.data, self.lines)
        # Для каждого потока свой лог. Потом сливаем в один и удаляем
        self.logging_dict = {}
        # Для сдвига окна при появлении
        self.thread_dict = {self.mode_description[i]['mode_name']: {} for i in self.mode_description}
        self.thread = None
        self.default_dict = {'mode_description': self.mode_description, 'logging_dict': self.logging_dict,
                             'thread_dict': self.thread_dict, 'default_path': self.default_path,
                             'all_doc': 0, 'now_doc': 0}

    def start_document(self, document):  # Запускаем окно с настройками по умолчанию.
        os.startfile(pathlib.Path(self.path_for_default, document))

    def sorting(self):  # Запускаем окно для сортировки.
        window_add = SortingFile(self, logging)
        window_add.exec_()

    def text_changed(self):  # Если изменился выбор принтера
        self.lineEdit_printer.setText(self.comboBox_printer.currentText())

    def insert_main(self):
        queue_main_insert = queue.Queue(maxsize=1)
        mode_name = self.mode_description['insert_main']['mode_name']
        name_dir = self.lineEdit_main_start_path_insert_dir.text().strip()
        out_dict = {
            'package': True if self.action_package.isChecked() else False,
            'action_mo': True if self.action_report_MO.isChecked() else False,
            'start_path': self.lineEdit_main_start_path_insert_dir.text().strip(),
            'finish_path': self.lineEdit_main_finish_path_insert_dir.text().strip(),
            'checkBox_file_num': True if self.checkBox_main_file_num.isChecked() else False,
            'file_num': self.lineEdit_main_file_num_path.text().strip(),
            'checkBox_signature': True if self.checkBox_main_signature.isChecked() else False,
            'path_signature': self.lineEdit_main_path_signature_dir.text().strip(),
            'radioButton_FSB': self.radioButton_main_group1_FSB_df.isChecked(),
            'radioButton_FSTEK': self.radioButton_main_group1_FSTEK_df.isChecked(),
            'classified': self.comboBox_main_classified.currentText().strip(),
            'num_scroll': self.lineEdit_main_num_scroll.text().strip(),
            'list_item': self.lineEdit_main_list_item.text().strip(),
            'checkBox_add_list_item': self.checkBox_main_add_list_item.isChecked(),
            'add_list_item': self.lineEdit_main_add_list_item.text().strip(),
            'number': self.lineEdit_main_secret_number.text().strip(),
            'hdd_number': self.lineEdit_main_HDD_number.text().strip(),
            'telephone': self.lineEdit_main_telephone.text().strip(),
            'conclusion': self.lineEdit_main_conclusion_executor.text().strip(),
            'checkBox_conclusion_number': self.checkBox_main_conclusion_number.isChecked(),
            'conclusion_number': self.lineEdit_main_conclusion_number.text().strip(),
            'conclusion_number_date': self.dateEdit_main_add_conclusion_date.date().toString('dd.MM.yyyy'),
            'protocol': self.lineEdit_main_protocol_executor.text().strip(),
            'prescription': self.lineEdit_main_prescription_executor.text().strip(),
            'executor_acc_sheet': self.lineEdit_main_acc_sheet_executor.text().strip(),
            'telephone_acc_sheet': self.lineEdit_main_telephone_acc_sheet_executor.text().strip(),
            'add_telephone': self.lineEdit_main_add_telephone.text().strip(),
            'print_executor': self.lineEdit_main_print_executor.text().strip(),
            'date': self.dateEdit_main_exec_date.date().toString('dd.MM.yyyy'),
            'inventory_executor': self.lineEdit_main_inventory_executor.text().strip(),
            'application_executor': self.lineEdit_main_application_executor.text().strip(),
            'act_executor': self.lineEdit_main_act_executor.text().strip(),
            'statement_executor': self.lineEdit_main_statement_executor.text().strip(),
            'inventory_insert': self.groupBox_main_inventory_insert.isChecked(),
            'radioButton_40_num': self.radioButton_main_group2_40_num.isChecked(),
            'radioButton_all_doc': self.radioButton_main_group2_all_doc.isChecked(),
            'flag_inventory': False,
            'account_position': self.lineEdit_main_account_position.text().strip(),
            'account_executor': self.lineEdit_main_account_executor.text().strip(),
            'account_path': self.lineEdit_main_account_path_dir.text().strip(),
            'form27_insert': self.groupBox_main_form27_insert.isChecked(),
            'form27_firm': self.lineEdit_main_form27_firm.text().strip(),
            'form27_path': self.lineEdit_main_form27_path_dir.text().strip(),
            'main_sp': self.groupBox_main_sp.isChecked(),
            'sp_path_dir': self.lineEdit_main_sp_path_dir.text().strip(),
            'sp_path_file': self.lineEdit_main_file_sp_path.text().strip(),
            'checkBox_gk': self.checkBox_main_name_gk.isChecked(),
            'name_gk': self.lineEdit_main_name_gk.text().strip(),
            'conclusion_sp': self.checkBox_main_conclusion_sp.isChecked(),
            'protocol_sp': self.checkBox_main_protocol_sp.isChecked(),
            'prescription_sp': self.checkBox_main_prescription_sp.isChecked(),
            'infocard_sp': self.checkBox_main_infocard_sp.isChecked(),
            'check_sp': [],
            'main_instance': self.groupBox_main_instance.isChecked(),
            'number_instance': self.lineEdit_main_number_instance.text().strip(),
            'conclusion_instance': self.checkBox_main_conclusion_instance.isChecked(),
            'protocol_instance': self.checkBox_main_protocol_instance.isChecked(),
            'prescription_instance': self.checkBox_main_prescription_instance.isChecked(),
        }
        data = {**self.default_dict, **out_dict,
                'queue': queue_main_insert, 'mode_name': mode_name, 'name_dir': name_dir,
                'start_function': format_doc}
        start_thread(data, self.logging_dict, self.thread_dict, self, doc_format, StartThreading)

    def print_main(self):
        queue_main_print = queue.Queue(maxsize=1)
        mode_name = self.mode_description['print_main']['mode_name']
        name_dir = self.lineEdit_main_start_path_print_dir.text().strip()
        out_dict = {
            'package': True if self.action_package.isChecked() else False,
            'start_path': self.lineEdit_main_start_path_print_dir.text().strip(),
            'path_account_num': self.lineEdit_main_file_account_numbers_path.text().strip(),
            'check_box_add_account_num': True if self.checkBox_main_add_account_numbers.isChecked() else False,
            'add_path_account_num': self.lineEdit_main_add_account_numbers_path_file.text().strip(),
            'name_printer': self.lineEdit_main_printer.text().strip(),
            'check_box_from_27': True if self.checkBox_main_file_form27.isChecked() else False,
            'path_form_27': self.lineEdit_main_path_file_form27_print.text().strip(),
            'print_order': True if self.checkBox_main_print_order.isChecked() else False,
            'fsb': True if self.radioButton_main_group3_FSB_print.isChecked() else False,
            'fstek': True if self.radioButton_main_group3_FSTEK_print.isChecked() else False,
            'service': '',
            'conclusion': True if self.checkBox_main_conclusion_print.isChecked() else False,
            'protocol': True if self.checkBox_main_protocol_print.isChecked() else False,
            'prescription': True if self.checkBox_main_prescription_print.isChecked() else False,
            'duplex': True if self.radioButton_main_group4_duplex.isChecked() else False,
            'last_duplex': True if self.radioButton_main_group4_last_duplex.isChecked() else False,
            'one_side': True if self.radioButton_main_group4_one_side.isChecked() else False,
        }
        data = {**self.default_dict, **out_dict,
                'queue': queue_main_print, 'mode_name': mode_name, 'name_dir': name_dir,
                'start_function': print_docs}
        start_thread(data, self.logging_dict, self.thread_dict, self, doc_print, StartThreading)

    def insert_41101(self):
        queue_41101_insert = queue.Queue(maxsize=1)
        mode_name = self.mode_description['insert_41101']['mode_name']
        name_dir = self.lineEdit_main_start_path_insert_dir.text().strip()
        out_dict = {
            'package': True if self.action_package.isChecked() else False,
            'action_mo': True if self.action_report_MO.isChecked() else False,
            'start_path': self.lineEdit_41101_start_path_insert_dir.text().strip(),
            'finish_path': self.lineEdit_41101_finish_path_insert_dir.text().strip(),
            'checkBox_file_num': True if self.checkBox_41101_file_num.isChecked() else False,
            'file_num': self.lineEdit_41101_file_num_path.text().strip(),
            'checkBox_signature': True if self.checkBox_41101_signature.isChecked() else False,
            'path_signature': self.lineEdit_41101_path_signature_dir.text().strip(),
            'radioButton_FSB': self.radioButton_41101_group5_FSB_df.isChecked(),
            'radioButton_FSTEK': self.radioButton_41101_group5_FSTEK_df.isChecked(),
            'classified': self.comboBox_41101_classified.currentText().strip(),
            'num_scroll': self.lineEdit_41101_num_scroll.text().strip(),
            'list_item': self.lineEdit_41101_list_item.text().strip(),
            'checkBox_add_list_item': self.checkBox_41101_add_list_item.isChecked(),
            'add_list_item': self.lineEdit_41101_add_list_item.text().strip(),
            'number': self.lineEdit_41101_secret_number.text().strip(),
            'hdd_number': self.lineEdit_41101_HDD_number.text().strip(),
            'telephone': self.lineEdit_41101_telephone.text().strip(),
            'conclusion': self.lineEdit_41101_conclusion_executor.text().strip(),
            'checkBox_conclusion_number': self.checkBox_41101_conclusion_number.isChecked(),
            'conclusion_number': self.lineEdit_41101_conclusion_number.text().strip(),
            'conclusion_number_date': self.dateEdit_41101_add_conclusion_date.date().toString('dd.MM.yyyy'),
            'protocol': self.lineEdit_41101_protocol_executor.text().strip(),
            'prescription': self.lineEdit_41101_prescription_executor.text().strip(),
            'executor_acc_sheet': self.lineEdit_41101_acc_sheet_executor.text().strip(),
            'telephone_acc_sheet': self.lineEdit_41101_telephone_acc_sheet_executor.text().strip(),
            'add_telephone': self.lineEdit_41101_add_telephone.text().strip(),
            'print_executor': self.lineEdit_41101_print_executor.text().strip(),
            'date': self.dateEdit_41101_exec_date.date().toString('dd.MM.yyyy'),
            'inventory_executor': self.lineEdit_41101_inventory_executor.text().strip(),
            'application_executor': self.lineEdit_41101_application_executor.text().strip(),
            'act_executor': self.lineEdit_41101_act_executor.text().strip(),
            'statement_executor': self.lineEdit_41101_statement_executor.text().strip(),
            'inventory_insert': self.groupBox_41101_inventory_insert.isChecked(),
            'radioButton_40_num': self.radioButton_41101_group6_40_num.isChecked(),
            'radioButton_all_doc': self.radioButton_41101_group6_all_doc.isChecked(),
            'flag_inventory': False,
            'account_position': self.lineEdit_41101_account_position.text().strip(),
            'account_executor': self.lineEdit_41101_account_executor.text().strip(),
            'account_path': self.lineEdit_41101_account_path_dir.text().strip(),
            'form27_insert': self.groupBox_41101_form27_insert.isChecked(),
            'form27_firm': self.lineEdit_41101_form27_firm.text().strip(),
            'form27_path': self.lineEdit_41101_form27_path_dir.text().strip(),
        }
        data = {**self.default_dict, **out_dict,
                'queue': queue_41101_insert, 'mode_name': mode_name, 'name_dir': name_dir,
                'start_function': format_doc}
        start_thread(data, self.logging_dict, self.thread_dict, self, doc_format, StartThreading)


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
