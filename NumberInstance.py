import os
import re
import shutil

import docx
from PyQt5.QtCore import QDir

import number_instance
from docx.shared import Pt

from PyQt5.QtWidgets import QDialog, QMessageBox, QFileDialog


class NumberInstance(QDialog, number_instance.Ui_Dialog):  # Для файла с номерами
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.pushButton_cancel.clicked.connect(lambda: self.close())
        self.pushButton_create.clicked.connect(self.create_file)
        self.pushButton_new.clicked.connect((lambda: self.browse(True)))
        self.pushButton_old.clicked.connect((lambda: self.browse(False)))

    def browse(self, num):  # Для кнопки открыть
        directory = QFileDialog.getExistingDirectory(self, "Find Files", QDir.currentPath())
        if directory:
            if num:
                self.lineEdit_path_new.setText(directory)
            else:
                self.lineEdit_path_old.setText(directory)

    def create_file(self):

        def check(n, e):
            for el_ in e:
                if n == el_:
                    return False
            return True

        path_old = self.lineEdit_path_old.text()
        if not path_old:
            QMessageBox.critical(self, 'УПС!', 'Путь к исходным файлам пуст')
            return
        if os.path.isdir(path_old):
            pass
        else:
            QMessageBox.critical(self, 'УПС!', 'Указанный путь к исходным файлам не является директорией')
            return
        path_new = self.lineEdit_path_new.text()
        if not path_new:
            QMessageBox.critical(self, 'УПС!', 'Путь к конечным файлам пуст')
            return
        if os.path.isdir(path_new):
            pass
        else:
            QMessageBox.critical(self, 'УПС!', 'Указанный путь к конечным файлам не является директорией')
            return
        number_instance_ = self.lineEdit_number_instance.text().strip()
        for i in number_instance_:
            if check(i, ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0', ' ', '-', ',', '.')):
                return ['УПС!', 'Есть лишние символы в номерах экземпляров']
        complect_num = number_instance_.replace(' ', '').replace(',', '.')
        if complect_num[0] == '.' or complect_num[0] == '-':
            return ['УПС!', 'Первый символ введён не верно']
        if complect_num[-1] == '.' or complect_num[-1] == '-':
            return ['УПС!', 'Последний символ введён не верно']
        for i in range(len(complect_num)):
            if complect_num[i] == '.' or complect_num[i] == '-':
                if complect_num[i + 1] == '.' or complect_num[i + 1] == '-':
                    return ['УПС!', 'Два разделителя номеров подряд']
        complect = []
        for element in complect_num.split('.'):
            if '-' in element:
                num1, num2 = int(element.partition('-')[0]), int(element.partition('-')[2])
                if num1 >= num2:
                    return ['УПС!', 'Диапазон номеров экземпляров указан не верно']
                else:
                    for el in range(num1, num2 + 1):
                        complect.append(el)
            else:
                complect.append(element)
        complect.sort()
        for number_folder in complect:
            os.mkdir(path_new + '\\' + str(number_folder) + ' экземпляр')
            for doc in os.listdir(path_old):
                shutil.copy2(path_old + '\\' + doc, path_new + '\\' + str(number_folder) + ' экземпляр' + '\\')
                doc_2 = docx.Document(os.path.abspath(path_new + '\\' + str(number_folder) +
                                                      ' экземпляр' + '\\' + doc))
                for p_2 in doc_2.sections[0].first_page_header.paragraphs:
                    if re.findall(r'№1', p_2.text):
                        text = re.sub(r'№1', '№' + str(number_folder), p_2.text)
                        p_2.text = text
                        for run in p_2.runs:
                            run.font.size = Pt(11)
                            run.font.name = 'Times New Roman'
                        break
                doc_2.save(os.path.abspath(path_new + '\\' + str(number_folder) +
                                           ' экземпляр' + '\\' + doc))  # Сохраняем
        self.close()  # Закрываем окно
