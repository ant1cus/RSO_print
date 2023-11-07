import itertools
import os
import pathlib
import re
import shutil
import traceback

import pandas as pd
from PyQt5.QtCore import QDir

import sorting

from PyQt5.QtWidgets import QDialog, QFileDialog, QMessageBox


class SortingFile(QDialog, sorting.Ui_Dialog):  # Для файла с номерами
    def __init__(self, parent, log):
        super().__init__()
        self.setupUi(self)
        self.parent = parent
        self.log = log
        self.pushButton_cancel.clicked.connect(lambda: self.close())
        self.pushButton_sorting.clicked.connect(self.sorting)
        self.pushButton_folder_document.clicked.connect((lambda: self.browse(self.lineEdit_path_folder_document)))
        self.pushButton_file_sp.clicked.connect((lambda: self.browse(self.lineEdit_path_file_sp)))
        self.pushButton_finish_folder.clicked.connect((lambda: self.browse(self.lineEdit_path_folder_finish)))

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

    def sorting(self):
        def check(n, e):
            for symbol in e:
                if n == symbol:
                    return True
            return False
        try:
            self.log.info('Старт сортировки')
            path_documents = self.lineEdit_path_folder_document.text().strip()
            if not path_documents:
                QMessageBox.critical(self, 'УПС!', 'Путь к папке с материалами СП пуст')
                return
            if os.path.isfile(path_documents):
                QMessageBox.critical(self, 'УПС!', 'Указанный путь к материалам СП не является директорией')
                return
            path_file = self.lineEdit_path_file_sp.text().strip()
            if not path_file:
                QMessageBox.critical(self, 'УПС!', 'Путь к файлу с номерами СП пуст')
                return
            if os.path.isdir(path_file):
                QMessageBox.critical(self, 'УПС!', 'Указанный путь к файлу с номерами СП не является файлом')
                return
            finish_path = self.lineEdit_path_folder_document.text().strip()
            if not finish_path:
                QMessageBox.critical(self, 'УПС!', 'Путь к конечной папке пуст')
                return
            if os.path.isfile(finish_path):
                QMessageBox.critical(self, 'УПС!', 'Указанный путь к конечной папке не является директорией')
                return
            name_gk = self.lineEdit_name_gk.text().strip()
            if name_gk is False:
                QMessageBox.critical(self, 'УПС!', 'Введите имя ГК')
                return
            for i in name_gk:
                if check(i, ("<", ">", ":", "/", "\\", "|", "?", "*", "\"")):
                    QMessageBox.critical(self, 'УПС!', 'Имя файла не должно содержать следующих символов:\n'
                                                       "<", ">", ":", "/", "\\", "|", "?", "*", "\"")
                    return
            check_file = [True if i.isChecked() else False for i in [self.checkBox_conclusion_sp,
                                                                     self.checkBox_protocol_sp,
                                                                     self.checkBox_preciption_sp,
                                                                     self.checkBox_infocard_sp]]
            if all(i is False for i in check_file):
                QMessageBox.critical(self, 'УПС!', 'Не выбран ни один документ для проверки СП')
                return
            self.log.info('Данные проверены')
            errors = []
            path_dir_document = pathlib.Path(finish_path, name_gk)
            os.makedirs(path_dir_document, exist_ok=True)
            df_number_sp = pd.read_excel(str(pathlib.Path(path_file)), sheet_name=0, header=None)
            df_number_sp.fillna(False, inplace=True)
            df_number_sp.drop(0, inplace=True)
            for name1, name2 in itertools.zip_longest(df_number_sp[0], df_number_sp[1]):
                if name1:
                    os.makedirs(str(pathlib.Path(path_dir_document, str(name1) + ' В')), exist_ok=True)
                if name2:
                    os.makedirs(str(pathlib.Path(path_dir_document, str(name2))), exist_ok=True)
            self.log.info('Созданы папки')
            files = [j_ for i_ in ['акт', 'заключение', 'протокол', 'предписание', 'инфокарта', 'result']
                     for j_ in os.listdir(path_documents) if re.findall(i_.lower(), j_.lower())]
            # result = [file for file in os.listdir(path_documents) if 'result' in file.lower()]
            self.log.info('Файлы отсортированы, перемещение')
            # if result:
            #     shutil.copy(str(pathlib.Path(path_documents, result[0])), str(pathlib.Path(path_dir_document)))
            for file in files:
                if 'акт' in file.lower() or 'result' in file.lower():
                    shutil.copy(str(pathlib.Path(path_documents, file)), str(pathlib.Path(path_dir_document)))
                else:
                    no_sn_in_sp = True
                    sn_number = file.rpartition(' ')[0].rpartition(' ')[2]
                    for folder_sp in os.listdir(str(path_dir_document)):
                        if re.findall(sn_number, folder_sp):
                            no_sn_in_sp = False
                            shutil.copy(str(pathlib.Path(path_documents, file)),
                                        str(pathlib.Path(path_dir_document, folder_sp)))
                    if no_sn_in_sp:
                        errors.append('Документ с с.н. ' + sn_number + ' (' + file + ') не найден в материалах СП')
            for folder_sp in os.listdir(str(pathlib.Path(path_dir_document))):
                if os.path.isdir(str(pathlib.Path(path_dir_document, folder_sp))):
                    file_sp = [file.partition(' ')[0].lower() for file in
                               os.listdir(str(pathlib.Path(path_dir_document, folder_sp)))]
                    for ind, check_file in enumerate(['заключение', 'протокол', 'предписание', 'инфокарта']):
                        if check_file[ind] and (check_file in file_sp) is False:
                            errors.append('В папке ' + str(folder_sp) + ' отсутствует ' + check_file)
            if errors:
                self.log.info('Есть ошибки')
                self.log.info('\n'.join(errors))
                self.parent.on_message_changed('УПС!', '\n'.join(errors))
            self.close()  # Закрываем окно
            self.log.info('Конец')
        except BaseException as error:
            QMessageBox.critical(self, 'УПС!', 'Что-то пошло не так')
            self.log.error("Ошибка:\n " + str(error) + '\n' + traceback.format_exc())
            return

