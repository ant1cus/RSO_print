import os
import queue
import sys
import pathlib
import logging

import win32print

import Main
import about

from general_function import browse, default_settings, default_data, rewrite_settings, start_thread
from create_number_instance import create_number_instance
from create_account_number import create_account_number
from SortingFile import SortingFile
from Check import (check_doc_format, check_doc_print, check_create_instance_number, check_create_account_number,
                   check_print_files, check_print_certification, check_word2pdf, check_sign2pdf)
from StartThread import StartThreading
from format_docs import format_doc
from print_docs import print_docs
from print_all_files import print_all_files
from print_certification import print_certification
from convert_word2pdf import convert_word2pdf
from sign2pdf import insert_sign2pdf

from PyQt5 import QtPrintSupport

from PyQt5.QtCore import (QTranslator, QLocale, QLibraryInfo, QObject)
from PyQt5.QtWidgets import (QMainWindow, QApplication, QDialog)


class AboutWindow(QDialog, about.Ui_Dialog):  # Для отображения информации
    def __init__(self):
        super().__init__()
        self.setupUi(self)


def about():  # Открываем окно с описанием
    window_add = AboutWindow()
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
                                 'create_instance_number': {'mode_name': 'create_instance_number',
                                                            'title': 'Создание номеров экземпляра в папке',
                                                            'cancel': 'Создание номеров экземпляра в папке'
                                                                      ' «name_dir» отменено пользователем',
                                                            'exception': 'Создание номеров экземпляра в папке'
                                                                         ' «name_dir» не завершено из-за ошибки',
                                                            'success': 'Создание номеров экземпляра в папке'
                                                                       ' «name_dir» успешно завершено',
                                                            'error': 'Создание номеров экземпляра в папке'
                                                                     ' «name_dir» завершено с ошибками'
                                                            },
                                 'create_account_number':  {'mode_name': 'create_account_number',
                                                            'title': 'Создание файла учетных номеров в папке',
                                                            'cancel': 'Создание файла учетных номеров в папке'
                                                                      ' «name_dir» отменено пользователем',
                                                            'exception': 'Создание файла учетных номеров в папке'
                                                                         ' «name_dir» не завершено из-за ошибки',
                                                            'success': 'Создание файла учетных номеров в папке'
                                                                       ' «name_dir» успешно завершено',
                                                            'error': 'Создание файла учетных номеров в папке'
                                                                     ' «name_dir» завершено с ошибками'
                                                            },
                                 'print_files': {'mode_name': 'print_files',
                                                 'title': 'Печать файлов в папке',
                                                 'cancel': 'Печать файлов в папке «name_dir» отменена пользователем',
                                                 'exception': 'Печать файлов в папке «name_dir»'
                                                              ' не завершена из-за ошибки',
                                                 'success': 'Печать файлов в папке «name_dir» успешно завершена',
                                                 'error': 'Печать файлов в папке «name_dir» завершена с ошибками'
                                                 },
                                 'print_certification': {'mode_name': 'print_certification',
                                                         'title': 'Печать лаборатории сертификации в папке',
                                                         'cancel': 'Печать лаборатории сертификации в папке «name_dir»'
                                                                   ' отменена пользователем',
                                                         'exception': 'Печать лаборатории сертификации в папке '
                                                                      '«name_dir» не завершена из-за ошибки',
                                                         'success': 'Печать лаборатории сертификации в папке «name_dir»'
                                                                    ' успешно завершена',
                                                         'error': 'Печать лаборатории сертификации в папке «name_dir» '
                                                                  'завершена с ошибками'
                                                         },
                                 'word2pdf': {'mode_name': 'word2pdf',
                                              'title': 'Преобразование Word в PDF в папке',
                                              'cancel': 'Преобразование Word в PDF в папке «name_dir»'
                                                        ' отменено пользователем',
                                              'exception': 'Преобразование Word в PDF в папке «name_dir»'
                                                           ' не завершено из-за ошибки',
                                              'success': 'Преобразование Word в PDF в папке «name_dir»'
                                                         ' успешно завершено',
                                              'error': 'Преобразование Word в PDF в папке «name_dir»'
                                                       ' завершено с ошибками'
                                              },
                                 'sign2pdf': {'mode_name': 'sign2pdf',
                                              'title': 'Вставить подпись в PDF в папке',
                                              'cancel': 'Вставка подписи в PDF в папке «name_dir»'
                                                        ' отменена пользователем',
                                              'exception': 'Вставка подписи в PDF в папке «name_dir»'
                                                           ' не завершена из-за ошибки',
                                              'success': 'Вставка подписи в PDF в папке «name_dir»'
                                                         ' успешно завершена',
                                              'error': 'Вставка подписи в PDF в папке «name_dir»'
                                                       ' завершена с ошибками'
                                              },
                                 }
        self.widget_name = {
            'insertMain': {'grid': 'gridLayout_insertMain', 'frame': 'groupBox_insertMain',
                           'action': 'action_insert_main', 'tab': 'insertMain'},
            'addInsertMain': {'grid': 'gridLayout_addInsertMain', 'frame': 'groupBox_addInsertMain',
                              'action': 'action_sorting_file', 'tab': 'sorting'},
            'printMain': {'grid': 'gridLayout_printMain', 'frame': 'groupBox_printMain',
                          'action': 'action_print_main', 'tab': 'printMain'},
            'insert41101': {'grid': 'gridLayout_insert41101', 'frame': 'groupBox_insert41101',
                            'action': 'action_insert_41101', 'tab': 'insert_41101'},
            'addInsert41101': {'grid': 'gridLayout_addInsert41101', 'frame': 'groupBox_addInsert41101',
                               'action': 'action_sorting_file', 'tab': 'sorting'},
            'print41101': {'grid': 'gridLayout_print41101', 'frame': 'groupBox_print41101',
                           'action': 'action_print_41101', 'tab': 'print_41101'},
            'module': {'grid': 'gridLayout_module', 'frame': 'groupBox_module',
                       'action': 'action_module', 'tab': 'module'}
        }
        self.pushButton_main_start_path_insert_dir.clicked.connect(
            lambda: browse(self, self.pushButton_main_start_path_insert_dir,
                           self.lineEdit_main_start_path_insert_dir, self.default_path))
        self.pushButton_main_finish_path_insert_dir.clicked.connect(
            lambda: browse(self, self.pushButton_main_finish_path_insert_dir,
                           self.lineEdit_main_finish_path_insert_dir, self.default_path))
        self.pushButton_main_file_num.clicked.connect(
            lambda: browse(self, self.pushButton_main_file_num,
                           self.lineEdit_main_file_num_path, self.default_path))
        self.pushButton_main_path_signature_dir.clicked.connect(
            lambda: browse(self, self.pushButton_main_path_signature_dir,
                           self.lineEdit_main_path_signature_dir, self.default_path))
        self.pushButton_main_account_path_dir.clicked.connect(
            lambda: browse(self, self.pushButton_main_account_path_dir,
                           self.lineEdit_main_account_path_dir, self.default_path))
        self.pushButton_main_form27_path_dir.clicked.connect(
            lambda: browse(self, self.pushButton_main_form27_path_dir,
                           self.lineEdit_main_form27_path_dir, self.default_path))
        self.pushButton_main_folder_sp_dir.clicked.connect(
            lambda: browse(self, self.pushButton_main_folder_sp_dir,
                           self.lineEdit_main_sp_path_dir, self.default_path))
        self.pushButton_main_file_sp.clicked.connect(
            lambda: browse(self, self.pushButton_main_file_sp,
                           self.lineEdit_main_file_sp_path, self.default_path))
        self.pushButton_main_start_path_print_dir.clicked.connect(
            lambda: browse(self, self.pushButton_main_start_path_print_dir,
                           self.lineEdit_main_start_path_print_dir, self.default_path))
        self.pushButton_main_file_form27_print.clicked.connect(
            lambda: browse(self, self.pushButton_main_file_form27_print,
                           self.lineEdit_main_path_file_form27_print, self.default_path))
        self.pushButton_main_file_account_numbers.clicked.connect(
            lambda: browse(self, self.pushButton_main_file_account_numbers,
                           self.lineEdit_main_file_account_numbers_path, self.default_path))
        self.pushButton_main_add_account_numbers.clicked.connect(
            lambda: browse(self, self.pushButton_main_add_account_numbers,
                           self.lineEdit_main_add_account_numbers_path, self.default_path))
        self.pushButton_41101_start_path_insert_dir.clicked.connect(
            lambda: browse(self, self.pushButton_41101_start_path_insert_dir,
                           self.lineEdit_41101_start_path_insert_dir, self.default_path))
        self.pushButton_41101_finish_path_insert_dir.clicked.connect(
            lambda: browse(self, self.pushButton_41101_finish_path_insert_dir,
                           self.lineEdit_41101_finish_path_insert_dir, self.default_path))
        self.pushButton_41101_file_num.clicked.connect(
            lambda: browse(self, self.pushButton_41101_file_num,
                           self.lineEdit_41101_file_num_path, self.default_path))
        self.pushButton_41101_account_path_dir.clicked.connect(
            lambda: browse(self, self.pushButton_41101_account_path_dir,
                           self.lineEdit_41101_account_path_dir, self.default_path))
        self.pushButton_41101_form27_path_dir.clicked.connect(
            lambda: browse(self, self.pushButton_41101_form27_path_dir,
                           self.lineEdit_41101_form27_path_dir, self.default_path))
        self.pushButton_41101_start_path_print_dir.clicked.connect(
            lambda: browse(self, self.pushButton_41101_start_path_print_dir,
                           self.lineEdit_41101_start_path_print_dir, self.default_path))
        self.pushButton_41101_file_form27_print.clicked.connect(
            lambda: browse(self, self.pushButton_41101_file_form27_print,
                           self.lineEdit_41101_path_file_form27_print, self.default_path))
        self.pushButton_41101_file_account_numbers.clicked.connect(
            lambda: browse(self, self.pushButton_41101_file_account_numbers,
                           self.lineEdit_41101_file_account_numbers_path, self.default_path))
        self.pushButton_41101_add_account_numbers.clicked.connect(
            lambda: browse(self, self.pushButton_41101_add_account_numbers,
                           self.lineEdit_41101_add_account_numbers_path, self.default_path))
        self.pushButton_module_path_account_dir.clicked.connect(
            lambda: browse(self, self.pushButton_module_path_account_dir,
                           self.lineEdit_module_path_account_finish_dir, self.default_path))
        self.pushButton_module_path_start_instance_dir.clicked.connect(
            lambda: browse(self, self.pushButton_module_path_start_instance_dir,
                           self.lineEdit_module_path_instance_start_dir, self.default_path))
        self.pushButton_module_path_finish_instance_dir.clicked.connect(
            lambda: browse(self, self.pushButton_module_path_finish_instance_dir,
                           self.lineEdit_module_path_account_finish_dir, self.default_path))
        self.pushButton_module_path_print_files_dir.clicked.connect(
            lambda: browse(self, self.pushButton_module_path_print_files_dir,
                           self.lineEdit_module_path_print_files_dir, self.default_path))
        self.pushButton_module_print_CL_start_path_print_file.clicked.connect(
            lambda: browse(self, self.pushButton_module_print_CL_start_path_print_file,
                           self.lineEdit_module_print_CL_start_path_print_file, self.default_path))
        self.pushButton_module_print_CL_start_path_print_dir.clicked.connect(
            lambda: browse(self, self.pushButton_module_print_CL_start_path_print_dir,
                           self.lineEdit_module_print_CL_start_path_print_file, self.default_path))
        self.pushButton_module_start_path_word2pdf_dir.clicked.connect(
            lambda: browse(self, self.pushButton_module_start_path_word2pdf_dir,
                           self.lineEdit_module_path_word2pdf_start_dir, self.default_path))
        self.pushButton_module_finish_path_word2pdf_dir.clicked.connect(
            lambda: browse(self, self.pushButton_module_finish_path_word2pdf_dir,
                           self.lineEdit_module_path_word2pdf_finish_dir, self.default_path))
        self.pushButton_module_path_sign2pdf_start_dir.clicked.connect(
            lambda: browse(self, self.pushButton_module_path_sign2pdf_start_dir,
                           self.lineEdit_module_path_sign2pdf_start_dir, self.default_path))
        self.pushButton_module_path_sign2pdf_finish_dir.clicked.connect(
            lambda: browse(self, self.pushButton_module_path_sign2pdf_finish_dir,
                           self.lineEdit_module_path_sign2pdf_finish_dir, self.default_path))
        self.pushButton_module_path_sign2pdf_signature_dir.clicked.connect(
            lambda: browse(self, self.pushButton_module_path_sign2pdf_signature_dir,
                           self.lineEdit_module_path_sign2pdf_signature_dir, self.default_path))
        # Для выбора принтера по умолчанию
        self.comboBox_main_printer.addItems(QtPrintSupport.QPrinterInfo.availablePrinterNames())
        self.comboBox_main_printer.currentTextChanged.connect(lambda: self.text_changed(self.lineEdit_main_printer,
                                                                                        self.comboBox_main_printer))
        self.lineEdit_main_printer.setText(QtPrintSupport.QPrinterInfo.defaultPrinterName())
        self.comboBox_41101_printer.addItems(QtPrintSupport.QPrinterInfo.availablePrinterNames())
        self.comboBox_41101_printer.currentTextChanged.connect(lambda: self.text_changed(self.lineEdit_41101_printer,
                                                                                         self.comboBox_41101_printer))
        self.lineEdit_41101_printer.setText(QtPrintSupport.QPrinterInfo.defaultPrinterName())
        self.comboBox_module_select_printer.addItems(QtPrintSupport.QPrinterInfo.availablePrinterNames())
        self.comboBox_module_select_printer.currentTextChanged.connect(
            lambda: self.text_changed(self.lineEdit_module_printer, self.comboBox_module_select_printer))
        self.lineEdit_module_printer.setText(QtPrintSupport.QPrinterInfo.defaultPrinterName())
        self.comboBox_module_print_CL_printer.addItems(QtPrintSupport.QPrinterInfo.availablePrinterNames())
        self.comboBox_module_print_CL_printer.currentTextChanged.connect(
            lambda: self.text_changed(self.lineEdit_module_print_CL_printer, self.comboBox_module_print_CL_printer))
        self.lineEdit_module_print_CL_printer.setText(QtPrintSupport.QPrinterInfo.defaultPrinterName())
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
                      'insertMain-checkBox_dont_check_number_files': ['Не сверять кол-во файлов',
                                                                      self.checkBox_main_dont_check_number_files],
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
                      'insert41101-checkBox_dont_check_number_files': ['Не сверять кол-во файлов',
                                                                       self.checkBox_41101_dont_check_number_files],
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
                                                                   self.lineEdit_41101_telephone_acc_sheet_executor],
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
                      'module-path_instance_start_dir': ['Путь к начальным экземплярам',
                                                         self.lineEdit_module_path_instance_start_dir],
                      'module-path_instance_finish_dir': ['Путь к конечным экземплярам',
                                                          self.lineEdit_module_path_instance_finish_dir],
                      'module-path_instance_number': ['Номера экземпляров', self.lineEdit_module_number_instance],
                      'module-path_account_finish_dir': ['Путь к новому файлу номеров',
                                                         self.lineEdit_module_path_account_finish_dir],
                      'module-account_number': ['Уч. номер, с', self.lineEdit_module_account_number],
                      'module-print_files': ['Исходные файлы для печати', self.lineEdit_module_path_print_files_dir],
                      'module-print_certification': ['Исходные файлы для печати сертификации',
                                                     self.lineEdit_module_print_CL_start_path_print_file],
                      'module-radioButton_group1': ['Метод печати сертификации',
                                                    [self.radioButton_module_group1_print_CL_duplex,
                                                     self.radioButton_module_group1_print_CL_last_duplex,
                                                     self.radioButton_module_group1_print_CL_one_side]],
                      'module-word2pdf_start_path': ['Начальная папка преобразования Word в PDF',
                                                     self.lineEdit_module_path_word2pdf_start_dir],
                      'module-word2pdf_finish_path': ['Конечная папка преобразования Word в PDF',
                                                      self.lineEdit_module_path_word2pdf_finish_dir],
                      'module-sign2pdf_start_path': ['Начальная папка вставки подписи в PDF',
                                                     self.lineEdit_module_path_sign2pdf_start_dir],
                      'module-sign2pdf_finish_path': ['Конечная папка вставки подписи в PDF',
                                                      self.lineEdit_module_path_sign2pdf_finish_dir],
                      'module-sign2pdf_signature_path': ['Папка с подписями для вставки в PDF',
                                                         self.lineEdit_module_path_sign2pdf_signature_dir],
                      }
        # Кнопки запуска
        self.pushButton_main_insert.clicked.connect(self.insert_main)
        self.pushButton_41101_insert.clicked.connect(self.insert_41101)
        self.pushButton_main_print.clicked.connect(self.print_main)
        self.pushButton_41101_print.clicked.connect(self.print_41101)
        self.pushButton_module_create_number_instance.clicked.connect(self.start_create_instance_number)
        self.pushButton_module_create_account_number.clicked.connect(self.start_create_account_number)
        self.pushButton_module_print_files.clicked.connect(self.print_files)
        self.pushButton_module_print_CL_print_files.clicked.connect(self.print_certification)
        self.pushButton_module_convert_word2pdf.clicked.connect(self.convert_word2pdf)
        self.pushButton_module_sign2pdf_create.clicked.connect(self.create_sign2pdf)
        # Кнопки в меню
        self.action_default.triggered.connect((lambda: default_settings(self, self.default_path,
                                                                        self.lines, self.widget_name)))
        self.action_about.triggered.connect(about)
        self.action_sorting.triggered.connect(self.sorting)
        self.action_instruction.triggered.connect(lambda: self.start_document('documents/Инструкция.docx'))
        self.action_registration.triggered.connect(lambda: self.start_document('documents/Номера для регистрации.xlsx'))
        self.action_sp.triggered.connect(lambda: self.start_document('documents/Номера СП.xlsx'))
        self.default_data = rewrite_settings(self.default_path)
        self.data = self.default_data["widget_settings"]
        if 'tab_order' in self.default_data['gui_settings']:
            self.tab_order = self.default_data['gui_settings']['tab_order']
        else:
            self.tab_order = {}
        if 'tab_visible' in self.default_data['gui_settings']:
            self.tab_visible = self.default_data['gui_settings']['tab_visible']
        else:
            self.tab_visible = {}
        default_data(self.data, self.lines)
        # Управление табами в виджете
        self.action_insert_main.triggered.connect(lambda: self.add_tab(self.action_insert_main))
        self.action_print_main.triggered.connect(lambda: self.add_tab(self.action_print_main))
        self.action_insert_41101.triggered.connect(lambda: self.add_tab(self.action_insert_41101))
        self.action_print_41101.triggered.connect(lambda: self.add_tab(self.action_print_41101))
        self.action_module.triggered.connect(lambda: self.add_tab(self.action_module))
        self.start_index = False
        self.start_name = False
        self.tabWidget.tabBar().tabMoved.connect(self.tab_)
        self.tabWidget.tabBarClicked.connect(self.tab_click)
        self.tabWidget.tabCloseRequested.connect(lambda index: self.tabWidget.removeTab(index))
        self.tab_for_paint = {}
        for tab in range(0, self.tabWidget.tabBar().count()):
            self.tab_for_paint[self.tabWidget.widget(tab).objectName()] = {}
            if self.tabWidget.widget(tab).objectName() not in self.tab_order.values():
                self.tab_order[str(len(self.tab_order))] = self.tabWidget.widget(tab).objectName()
                rewrite_settings(self.default_path, self.tab_order, 'tab_order')
                self.tab_visible[str(self.tabWidget.widget(tab).objectName())] = True
                rewrite_settings(self.default_path, self.tab_visible, 'tab_visible')
            self.tab_for_paint[self.tabWidget.widget(tab).objectName()]['widget'] = self.tabWidget.widget(tab)
            self.tab_for_paint[self.tabWidget.widget(tab).objectName()]['name'] = self.tabWidget.tabText(tab)
        self.tabWidget.clear()
        for tab in self.tab_order:
            action = self.findChild(QObject, self.widget_name[self.tab_order[tab]]['action'])
            if self.tab_visible[self.tab_order[tab]]:
                action.setChecked(True)
                self.tabWidget.addTab(self.tab_for_paint[self.tab_order[tab]]['widget'],
                                      self.tab_for_paint[self.tab_order[tab]]['name'])
            else:
                action.setChecked(False)
        self.tabWidget.tabBar().setCurrentIndex(0)
        # Для каждого потока свой лог. Потом сливаем в один и удаляем
        self.logging_dict = {}
        # Для сдвига окна при появлении
        self.thread_dict = {self.mode_description[i]['mode_name']: {} for i in self.mode_description}
        self.thread = None
        self.default_dict = {'mode_description': self.mode_description, 'logging_dict': self.logging_dict,
                             'thread_dict': self.thread_dict, 'default_path': self.default_path,
                             'all_doc': 0, 'now_doc': 0}

    def tab_(self, index):
        for tab in self.tab_order.items():
            if tab[1] == self.start_name and tab[1] == self.tabWidget.currentWidget().objectName():
                self.tab_order[str(index)], self.tab_order[tab[0]] = self.tab_order[tab[0]], self.tab_order[str(index)]
                break
            elif tab[1] == self.tabWidget.currentWidget().objectName():
                self.tab_order[str(index)], self.tab_order[tab[0]] = self.tab_order[tab[0]], self.tab_order[str(index)]
                break
        rewrite_settings(self.default_path, self.tab_order, 'tab_order')

    def tab_click(self, index):
        try:
            self.start_name = self.tab_order[str(index)]
        except KeyError:
            pass

    def add_tab(self, widget_action):
        name_open_tab = {self.tabWidget.widget(ind).objectName(): ind for ind
                         in range(0, self.tabWidget.tabBar().count())}
        for tab in self.widget_name:
            action = self.findChild(QObject, self.widget_name[tab]['action'])
            if action == widget_action:
                if action.isChecked():
                    if tab not in name_open_tab:
                        self.tabWidget.addTab(self.tab_for_paint[tab]['widget'],
                                              self.tab_for_paint[tab]['name'])
                    if self.tab_visible[tab] is False:
                        self.tab_visible[tab] = True
                        rewrite_settings(self.default_path, self.tab_visible, 'tab_visible')
                else:
                    if self.tab_visible[tab]:
                        self.tab_visible[tab] = False
                        rewrite_settings(self.default_path, self.tab_visible, 'tab_visible')

    def start_document(self, document):  # Запускаем окно с настройками по умолчанию.
        os.startfile(pathlib.Path(self.default_path, document))

    def sorting(self):  # Запускаем окно для сортировки.
        window_add = SortingFile(self, logging)
        window_add.exec_()

    def text_changed(self, line_edit, combo_box):  # Если изменился выбор принтера
        line_edit.setText(combo_box.currentText())
        win32print.SetDefaultPrinter(combo_box.currentText())

    def create_sign2pdf(self):
        queue_create_sign2pdf = queue.Queue(maxsize=1)
        mode_name = self.mode_description['sign2pdf']['mode_name']
        name_dir = self.lineEdit_module_path_sign2pdf_start_dir.text().strip()
        out_dict = {
            'start_path': self.lineEdit_module_path_sign2pdf_start_dir.text().strip(),
            'finish_path': self.lineEdit_module_path_sign2pdf_finish_dir.text().strip(),
            'signature_path': self.lineEdit_module_path_sign2pdf_signature_dir.text().strip(),
        }
        data = {**self.default_dict, **out_dict,
                'queue': queue_create_sign2pdf, 'mode_name': mode_name, 'name_dir': name_dir,
                'start_function': insert_sign2pdf}
        start_thread(data, self.logging_dict, self.thread_dict, self, check_sign2pdf, StartThreading)

    def convert_word2pdf(self):
        queue_convert_word2pdf = queue.Queue(maxsize=1)
        mode_name = self.mode_description['word2pdf']['mode_name']
        name_dir = self.lineEdit_module_path_word2pdf_start_dir.text().strip()
        out_dict = {
            'start_path': self.lineEdit_module_path_word2pdf_start_dir.text().strip(),
            'finish_path': self.lineEdit_module_path_word2pdf_finish_dir.text().strip(),
        }
        data = {**self.default_dict, **out_dict,
                'queue': queue_convert_word2pdf, 'mode_name': mode_name, 'name_dir': name_dir,
                'start_function': convert_word2pdf}
        start_thread(data, self.logging_dict, self.thread_dict, self, check_word2pdf, StartThreading)

    def print_certification(self):
        queue_print_certification = queue.Queue(maxsize=1)
        mode_name = self.mode_description['print_certification']['mode_name']
        name_dir = self.lineEdit_module_print_CL_start_path_print_file.text().strip()
        name_dir = pathlib.Path(name_dir).parent if pathlib.Path(name_dir).is_file() else name_dir
        out_dict = {
            'start_path': self.lineEdit_module_print_CL_start_path_print_file.text().strip(),
            'name_printer': self.lineEdit_module_print_CL_printer.text().strip(),
            'duplex': True if self.radioButton_module_group1_print_CL_duplex.isChecked() else False,
            'last_duplex': True if self.radioButton_module_group1_print_CL_last_duplex.isChecked() else False,
            'one_side': True if self.radioButton_module_group1_print_CL_one_side.isChecked() else False,
        }
        data = {**self.default_dict, **out_dict,
                'queue': queue_print_certification, 'mode_name': mode_name, 'name_dir': name_dir,
                'start_function': print_certification}
        start_thread(data, self.logging_dict, self.thread_dict, self, check_print_certification, StartThreading)

    def print_files(self):
        queue_print_files = queue.Queue(maxsize=1)
        mode_name = self.mode_description['print_files']['mode_name']
        name_dir = self.lineEdit_module_path_print_files_dir.text().strip()
        out_dict = {
            'start_path': self.lineEdit_module_path_print_files_dir.text().strip(),
            'printer': self.lineEdit_module_printer.text().strip()
        }
        data = {**self.default_dict, **out_dict,
                'queue': queue_print_files, 'mode_name': mode_name, 'name_dir': name_dir,
                'start_function': print_all_files}
        start_thread(data, self.logging_dict, self.thread_dict, self, check_print_files, StartThreading)

    def start_create_instance_number(self):
        queue_create_instance_number = queue.Queue(maxsize=1)
        mode_name = self.mode_description['create_instance_number']['mode_name']
        name_dir = self.lineEdit_module_path_instance_start_dir.text().strip()
        out_dict = {
            'start_path': self.lineEdit_module_path_instance_start_dir.text().strip(),
            'finish_path': self.lineEdit_module_path_instance_finish_dir.text().strip(),
            'number_instance': self.lineEdit_module_number_instance.text().strip()
        }
        data = {**self.default_dict, **out_dict,
                'queue': queue_create_instance_number, 'mode_name': mode_name, 'name_dir': name_dir,
                'start_function': create_number_instance}
        start_thread(data, self.logging_dict, self.thread_dict, self, check_create_instance_number, StartThreading)

    def start_create_account_number(self):
        queue_create_account_number = queue.Queue(maxsize=1)
        mode_name = self.mode_description['create_account_number']['mode_name']
        name_dir = self.lineEdit_module_path_account_finish_dir.text().strip()
        out_dict = {
            'start_path': self.lineEdit_module_path_account_finish_dir.text().strip(),
            'account_number': self.lineEdit_module_account_number.text().strip()
        }
        data = {**self.default_dict, **out_dict,
                'queue': queue_create_account_number, 'mode_name': mode_name, 'name_dir': name_dir,
                'start_function': create_account_number}
        start_thread(data, self.logging_dict, self.thread_dict, self, check_create_account_number, StartThreading)

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
            'dont_check': self.checkBox_main_dont_check_number_files.isChecked(),
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
        start_thread(data, self.logging_dict, self.thread_dict, self, check_doc_format, StartThreading)

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
            'check_box_from_27': True if self.checkBox_main_file_form27.isChecked() else False,
            'path_form_27': self.lineEdit_main_path_file_form27_print.text().strip(),
            'name_printer': self.lineEdit_main_printer.text().strip(),
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
        start_thread(data, self.logging_dict, self.thread_dict, self, check_doc_print, StartThreading)

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
            'dont_check': self.checkBox_41101_dont_check_number_files.isChecked(),
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
        start_thread(data, self.logging_dict, self.thread_dict, self, check_doc_format, StartThreading)

    def print_41101(self):
        queue_41101_print = queue.Queue(maxsize=1)
        mode_name = self.mode_description['print_41101']['mode_name']
        name_dir = self.lineEdit_41101_start_path_print_dir.text().strip()
        out_dict = {
            'package': True if self.action_package.isChecked() else False,
            'start_path': self.lineEdit_41101_start_path_print_dir.text().strip(),
            'path_account_num': self.lineEdit_41101_file_account_numbers_path.text().strip(),
            'check_box_add_account_num': True if self.checkBox_41101_add_account_numbers.isChecked() else False,
            'add_path_account_num': self.lineEdit_41101_add_account_numbers_path_file.text().strip(),
            'check_box_from_27': True if self.checkBox_41101_file_form27.isChecked() else False,
            'path_form_27': self.lineEdit_41101_path_file_form27_print.text().strip(),
            'name_printer': self.lineEdit_41101_printer.text().strip(),
            'print_order': True,
            'fsb': True if self.radioButton_41101_group7_FSB_print.isChecked() else False,
            'fstek': True if self.radioButton_41101_group7_FSTEK_print.isChecked() else False,
            'service': '',
            'conclusion': True if self.checkBox_41101_conclusion_print.isChecked() else False,
            'protocol': True if self.checkBox_41101_protocol_print.isChecked() else False,
            'prescription': True if self.checkBox_41101_prescription_print.isChecked() else False,
            'duplex': True if self.radioButton_41101_group8_duplex.isChecked() else False,
            'last_duplex': True if self.radioButton_41101_group8_last_duplex.isChecked() else False,
            'one_side': True if self.radioButton_41101_group8_one_side.isChecked() else False,
        }
        data = {**self.default_dict, **out_dict,
                'queue': queue_41101_print, 'mode_name': mode_name, 'name_dir': name_dir,
                'start_function': print_docs}
        start_thread(data, self.logging_dict, self.thread_dict, self, check_doc_print, StartThreading)


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
