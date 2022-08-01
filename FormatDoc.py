import datetime
import os
import pathlib
import re
import shutil
import time
import traceback
import zipfile

import docx
import docx2pdf
import fitz
import numpy
import openpyxl
import pythoncom
import pandas as pd
import openpyxl.styles
from PyQt5.QtCore import QThread, pyqtSignal
from docx.enum.section import WD_ORIENTATION
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
from openpyxl.utils import get_column_letter
from natsort import natsorted


class FormatDoc(QThread):  # Если требуется вставить колонтитулы
    progress = pyqtSignal(int)  # Сигнал для прогресс бара
    status = pyqtSignal(str)  # Сигнал для статус бара
    messageChanged = pyqtSignal(str, str)

    def __init__(self, output):  # Список переданных элементов.
        QThread.__init__(self)
        self.path_old = output[0]
        self.path_new = output[1]
        self.file_num = output[2]
        self.classified = output[3]
        self.num_scroll = output[4]
        self.list_item = output[5]
        self.number = output[6]
        self.executor = output[7]
        self.conclusion = output[8]
        self.prescription = output[9]
        self.print_people = output[10]
        self.date = output[11]
        self.executor_acc_sheet = output[12]
        self.account = output[13]
        self.flag_inventory = output[14]
        self.account_post = output[15]
        self.account_signature = output[16]
        self.account_path = output[17]
        self.firm = output[18]
        self.path_form_27 = output[19]
        self.second_copy = output[20]
        self.service = output[21]
        self.hdd_number = output[22]
        self.q = output[23]
        self.logging = output[24]
        self.package = output[25]
        self.num_1 = self.num_2 = 0

    def run(self):

        def format_doc_(path_old_, classified, list_item, num_scroll, account,
                        firm, logging, status, path_, file_num, num_1, num_2,
                        date, conclusion, executor, prescription, hdd_number,
                        print_people, progress, flag_inventory,
                        account_post, account_signature, account_path, executor_acc_sheet, service, path_form_27):
            def cell_write(style_for_doc, text_for_insert, number_rows=0):  # Заполнение ячеек в таблице в описи
                cells = table.rows[number_rows].cells  # Номер строки
                number_col = 0  # Номер столбца
                for elem in text_for_insert:
                    cells[number_col].text = elem  # Заполняем елемент
                    cells[number_col].paragraphs[
                        0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # Выравнивание по центру
                    cells[number_col].paragraphs[0].style = style_for_doc
                    if number_col == 1:  # Размер если ячейка с именем документа
                        cells[number_col].width = 12801600  # 1.4 * 914400
                    elif number_col == 3:  # Размер если ячейка с номером и грифом
                        cells[number_col].width = 10972800  # 1.2 * 914400
                    number_col += 1

            def rm(folder_path):
                try:
                    while len(os.listdir(folder_path)) != 0:
                        time.sleep(0.5)
                        for file_object in os.listdir(folder_path):
                            flag_ = True
                            while flag_:
                                try:
                                    file_object_path = os.path.join(folder_path, file_object)
                                    try:
                                        if os.path.isfile(file_object_path) or os.path.islink(file_object_path):
                                            os.remove(file_object_path)
                                        else:
                                            try:
                                                shutil.rmtree(file_object_path)
                                            except FileNotFoundError:
                                                pass
                                    except OSError:
                                        os.remove(folder_path)
                                    flag_ = False
                                except BaseException:
                                    pass
                except NotADirectoryError:
                    os.remove(folder_path)
                time.sleep(0.05)
                shutil.rmtree(folder_path)

            def pages_count(count_file, file_path):
                # Конвертируем
                while True:
                    try:
                        print(count_file)
                        print(file_path)
                        pythoncom.CoInitializeEx(0)
                        name_file_pdf = count_file + '.pdf'
                        self.logging.info('Конвертируем в пдф ' + count_file)
                        docx2pdf.convert(file_path + '\\' + count_file, file_path + '\\' + name_file_pdf)
                        input_file_pdf = fitz.open(file_path + '\\' + name_file_pdf)  # Открываем пдф
                        count_page = input_file_pdf.page_count  # Получаем кол-во страниц
                        input_file_pdf.close()  # Закрываем
                        self.logging.info('Удаляем пдф ' + count_file)
                        os.remove(file_path + '\\' + name_file_pdf)  # Удаляем пдф документ
                        self.logging.info('Вставляем страницы в ворд ' + count_file)
                        temp_docx = os.path.join(file_path, count_file)
                        temp_zip = os.path.join(file_path, count_file + ".zip")
                        temp_folder = os.path.join(file_path, "template")

                        if os.path.exists(temp_zip):
                            rm(temp_zip)
                        if os.path.exists(temp_folder):
                            rm(temp_folder)
                        if os.path.exists(file_path + '\\zip'):
                            rm(file_path + '\\zip')
                        os.rename(temp_docx, temp_zip)
                        os.mkdir(file_path + '\\zip')
                        with zipfile.ZipFile(temp_zip) as my_document:
                            my_document.extractall(temp_folder)
                        pages_xml = os.path.join(temp_folder, "docProps", "app.xml")
                        string = open(pages_xml, 'r', encoding='utf-8').read()
                        string = re.sub(r"<Pages>(\w*)</Pages>",
                                        "<Pages>" + str(count_page) + "</Pages>", string)
                        with open(pages_xml, "wb") as file_wb:
                            file_wb.write(string.encode("UTF-8"))
                        self.logging.info('Получае ворд из зип ' + count_file)
                        os.remove(temp_zip)
                        shutil.make_archive(temp_zip.replace(".zip", ""), 'zip', temp_folder)
                        os.rename(temp_zip, temp_docx)  # rename zip file to docx
                        rm(temp_folder)
                        rm(file_path + '\\zip')
                        return count_page
                    except BaseException as exept:
                        self.logging.error("Ошибка:\n " + str(exept) + '\n' + traceback.format_exc())

            def insert_header(doc_, pt_count, text_first_header_, text_for_foot_, hdd_number_, exec_,
                              print_people_, date_, path_new, name_file_, fso_):
                style_ = doc_.styles['Normal']
                font_ = style_.font
                font_.name = 'TimesNewRoman'
                font_.size = Pt(pt_count)
                header_ = doc_.sections[0].first_page_header  # Верхний колонтитул первой страницы
                head_1 = header_.paragraphs[0]  # Параграф
                head_1.insert_paragraph_before(text_first_header_)  # Вставялем перед колонитулом
                head_1 = header_.paragraphs[0]  # Выбираем новый первый параграф
                head_1.style = doc_.styles['Normal']
                head_1_format = head_1.paragraph_format  # Настройки параграфа
                head_1_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT  # Выравниваем по правому краю
                footer_ = doc_.sections[0].first_page_footer  # Нижний колонтитул первой страницы
                foot_ = footer_.paragraphs[0]  # Параграф
                foot_.text = text_for_foot_  # Текст
                foot_format_ = foot_.paragraph_format  # Настройки параграфа
                foot_format_.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT  # Выравнивание по левому краю
                doc_.sections[0].footer.paragraphs[0].text = text_for_foot_  # Номера для страниц
                # Выравниваем по левому краю
                doc_.sections[0].footer.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                doc_.add_section()  # Добавляем последнюю страницу
                last_ = doc_.sections[
                    len(doc_.sections) - 1].first_page_header  # Колонтитул для последней страницы
                last_.is_linked_to_previous = False  # Отвязываем от предыдущей секции чтобы не повторялись
                foot_ = doc_.sections[len(doc_.sections) - 1].first_page_footer  # Нижний колонтитул
                foot_.is_linked_to_previous = False  # Отвязываем
                # Текст для фонарика
                foot_.paragraphs[0].text = "Уч. № " + text_for_foot_ + \
                                           "\nОтп. 1 экз. в адрес\nс НЖМД, уч. № ЖД - " + hdd_number_ + \
                                           "\nИсп. " + exec_ + "\nПеч. " + print_people_ + "\n" + \
                                           date_ + "\nб/ч"
                foot_.paragraphs[0].style = doc.styles['Normal']
                if fso_:
                    if 'заключение' in name_file_.lower():
                        path_new = path_new + '\\' + 'Материалы по специальной проверке технических средств'
                    else:
                        path_new = path_new + '\\' + 'Материалы по специальным исследованиям технических средств'
                    try:
                        os.mkdir(path_new)
                    except FileExistsError:
                        pass
                doc_.save(os.path.abspath(path_new + '\\' + name_file_))  # Сохраняем

            def change_date(docum, param):
                RU_MONTH_VALUES = {
                    1: 'января',
                    2: 'февраля',
                    3: 'марта',
                    4: 'апреля',
                    5: 'мая',
                    6: 'июня',
                    7: 'июля',
                    8: 'августа',
                    9: 'сентября',
                    10: 'октября',
                    11: 'ноября',
                    12: 'декабря'
                }
                for parag_ in docum.paragraphs:
                    if re.findall(r'date', parag_.text):
                        text_date = re.sub(r'date', "«{}» {} {} г.".format(date.partition('.')[0],
                                                                           RU_MONTH_VALUES[int(date.partition('.')[2].partition('.')[0])],
                                                                           date.rpartition('.')[2]),
                                           parag_.text)
                        parag_.text = text_date
                        parag_.style = doc.styles['Normal']
                        for runs_ in parag_.runs:
                            runs_.font.size = Pt(12) if param else Pt(pt_num)
                        break

            os.chdir(path_old_)  # Меняем рабочую директорию
            # Параграф для колонтитула первой страницы
            text_first_header = classified + '\n' + list_item + '\nЭкз. №' + num_scroll
            fso = False
            for folder_ in os.listdir():
                if os.path.isdir(folder_):
                    fso = True
                    break
            if fso:
                docs = {}
                docs_not = {}
                for folder_ in os.listdir():
                    if 'проверке' in folder_.lower():
                        if folder_ != 'Материалы по специальной проверке технических средств':
                            self.messageChanged.emit('УПС!', 'Название папки «Материалы по специальной проверке'
                                                             ' технических средств» написано с ошибками')
                            return
                        docs[os.listdir(path_old_ + '\\' + folder_)[0]] = path_old_ + '\\' + folder_
                    elif 'исследованиям' in folder_.lower():
                        if folder_ != 'Материалы по специальным исследованиям технических средств':
                            self.messageChanged.emit('УПС!', 'Название папки «Материалы по специальным исследованиям'
                                                             ' технических средств» написано с ошибками')
                            return
                        file = os.listdir(path_old_ + '\\' + folder_)
                        for name_element in ['протокол', 'предписание']:
                            for element in file:
                                if name_element in element.lower():
                                    docs[element] = path_old_ + '\\' + folder_
                    elif 'дополнительные' in folder_.lower() and os.path.isdir(folder_):
                        if folder_ != 'Дополнительные материалы':
                            self.messageChanged.emit('УПС!', 'Название папки «Дополнительные материалы»'
                                                             ' написано с ошибками')
                            return
                        shutil.copytree(os.getcwd() + '\\' + folder_, path_ + '\\' + folder_)
                for element in os.listdir():
                    if 'сопроводит' in element.lower():
                        docs[element] = path_old_
                docx_for_progress = len(docs)

            else:
                docs = [file for file in os.listdir() if file[-4:] == 'docx']  # Список документов

                def sort(input_str):  # Ф-я для сортировка

                    try:
                        return float(input_str.partition('.')[2][:-5])
                    except ValueError:
                        return 1

                docs.sort(key=sort)  # Сортировка
                docs = natsorted(docs)
                docs_ = [j_ for i_ in
                         ['Заключение', 'Протокол', 'Приложение', 'Предписание', 'Форма 3', 'Опись',
                          'Сопроводит'] for j_ in docs if
                         re.findall(i_.lower(), j_.lower())]
                docs_not = [i_ for i_ in docs if i_ not in docs_ and '~' not in i_]
                docs = docs_not + docs_
                logging.info("Отсортированы документы:\n" + '\n'.join([i_ for i_ in docs]))
                for el in docs:
                    if el.endswith('.doc'):
                        status.emit('Документ ' + os.path.basename(el) + ' формата .doc'
                                                                         ' (необходим .docx). Замените файл')
                        return
                # Процент для прогресса
                docx_for_progress = 0
                for name_file in os.listdir():
                    if re.findall(r'приложение', name_file.lower()):
                        pass
                    else:
                        docx_for_progress += 1
            per = 90 if account else 100
            percent = (per - 10) / docx_for_progress if firm else per / docx_for_progress
            percent_val = 0
            conclusion_num = {}
            protocol = {}
            dict_40 = []  # Словарь для описи
            for_27 = []
            accompanying_doc = ''  # Проверка на сопровод
            exec_people = ''  # Для исполнителя документов
            text_for_foot = ''
            if self.second_copy:
                os.mkdir(path_ + '\\2 экземпляр')
            for el_ in docs:  # Для файлов в папке
                name_el = el_
                if type(docs) is dict:
                    os.chdir(docs[el_])
                logging.info("Преобразуем " + name_el)
                if name_el.lower() == 'форма 3.docx':
                    continue
                elif re.findall('сопроводит', name_el.lower()):
                    accompanying_doc = el_
                    continue
                pythoncom.CoInitializeEx(0)
                status.emit('Форматируем документ ' + name_el)
                doc = docx.Document(el_)  # Открываем
                if re.findall(r'приложение', name_el.lower()):
                    number_protocol = name_el.rpartition(' ')[2].rpartition('.')[0]
                    for appendix_num in for_27:
                        if re.findall('протокол', appendix_num[4].lower()) \
                                and re.findall(number_protocol, appendix_num[4]):
                            for p in doc.paragraphs:
                                if re.findall(r'\[ПРОТНОМ]', p.text):
                                    text = re.sub(r'\[ПРОТНОМ]', 'к протоколу уч. № ' + appendix_num[0] + ' от ' + date,
                                                  p.text)
                                    p.text = text
                                    p.style = doc.styles['Normal']
                                    for run in p.runs:
                                        run.font.size = Pt(pt_num)
                                    break
                            break
                    doc.save(os.path.abspath(path_ + '\\' + name_el))  # Сохраняем
                else:
                    if file_num:  # Если есть файл номеров
                        text_for_foot = dict_file[name_el.rpartition('.')[0]][0]  # Текст для нижнего колонтитула
                        date = dict_file[name_el.rpartition('.')[0]][1]  # Дата
                    else:
                        text_for_foot = num_1 + num_2 + 'c'  # Текст для нижнего колонитула
                    if re.findall(r'заключение'.lower(), name_el.lower()):
                        name_conclusion = name_el.rpartition('.')[0].rpartition(' ')[0]
                        conclusion_num[name_el] = text_for_foot
                        exec_people = conclusion
                        change_date(doc, True)
                    elif re.findall(r'протокол', name_el.lower()):
                        name_protocol = name_el.rpartition('.')[0].rpartition(' ')[0]
                        protocol[name_el] = text_for_foot
                        if len(conclusion_num) == 1:
                            conclusion_num_text = 'Уч. № ' + str(conclusion_num[list(conclusion_num.keys())[0]]) \
                                                  + ' от ' + date
                        else:
                            x = name_el.rpartition('.')[0].partition(' ')[2]
                            conclusion_num_text = 'Уч. № ' + str(conclusion_num[name_conclusion + ' ' + x + '.docx']) \
                                                  + ' от ' + date
                        for val_p, p in enumerate(doc.paragraphs):
                            if re.findall(r'\[ЗАКЛНОМ]', p.text):
                                text = re.sub(r'\[ЗАКЛНОМ]', conclusion_num_text, p.text)
                                p.text = text
                                p.style = doc.styles['Normal']
                                doc.paragraphs[val_p].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                                for run in p.runs:
                                    run.font.size = Pt(pt_num)
                                break
                        exec_people = executor
                        change_date(doc, False)
                    elif re.findall(r'предписание', name_el.lower()):
                        x = name_el.rpartition('.')[0].partition(' ')[2]
                        protocol_num_text = 'Уч. № ' + str(protocol[name_protocol + ' ' + x + '.docx']) + \
                                            ' от ' + date
                        if len(conclusion_num) == 1:
                            conclusion_num_text = 'Уч. № ' + str(conclusion_num[list(conclusion_num.keys())[0]]) + \
                                                  ' от ' + date
                        else:
                            conclusion_num_text = 'Уч. № ' + str(conclusion_num[name_conclusion + ' ' + x + '.docx']) + \
                                                  ' от ' + date
                        break_flag = 0
                        for val_p, p in enumerate(doc.paragraphs):
                            if re.findall(r'\[ЗАКЛНОМ]', p.text):
                                text = re.sub(r'\[ЗАКЛНОМ]', conclusion_num_text, p.text)
                                p.text = text
                                p.style = doc.styles['Normal']
                                doc.paragraphs[val_p].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                                for run in p.runs:
                                    run.font.size = Pt(pt_num)
                                break_flag += 1
                            if re.findall(r'\[ПРОТНОМ]', p.text):
                                text = re.sub(r'\[ПРОТНОМ]', protocol_num_text, p.text)
                                p.text = text
                                p.style = doc.styles['Normal']
                                doc.paragraphs[val_p].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                                for run in p.runs:
                                    run.font.size = Pt(pt_num)
                                break_flag += 1
                            if break_flag == 2:
                                break
                        exec_people = prescription
                        change_date(doc, False)
                    logging.info("Вставляем колонтитулы")
                    insert_header(doc, 11, text_first_header, text_for_foot, hdd_number,
                                  exec_people, print_people, date, path_, name_el, fso)
                    logging.info("Определяем количество страниц")
                    if fso:
                        path_to_file = path_ + '\\' + docs[el_].rpartition('\\')[2]
                        num_pages = pages_count(name_el, path_to_file)
                    else:
                        num_pages = pages_count(name_el, path_)
                    for_27.append([text_for_foot, date, classified, firm, name_el[:-5], executor, '1',
                                   '№1', str(num_pages - 1)])
                    if account:  # Если активирована опись добавляем
                        dict_40.append({name_el: [classified, text_for_foot, num_scroll, str(num_pages - 1)]})
                    if not file_num:  # Если нет файла номеров
                        num_2 = str(int(num_2) + 1)  # Увеличиваем значение для учетного номера
                    if self.second_copy:
                        for index, (value1, value2) in enumerate(zip(self.second_copy, ['заключение',
                                                                                        'протокол', 'предписание'])):
                            if value1 and re.findall(value2, el_.lower()):
                                for_27.append([False, False, False, False, name_el[:-5], False, '2',
                                               False, str(num_pages - 1)])
                                shutil.copy2(path_ + '\\' + el_, path_ + '\\2 экземпляр')
                                doc_2 = docx.Document(os.path.abspath(path_ + '\\2 экземпляр' + '\\' + el_))
                                for p_2 in doc_2.sections[0].first_page_header.paragraphs:
                                    if re.findall(r'№1', p_2.text):
                                        text = re.sub(r'№1', '№2', p_2.text)
                                        p_2.text = text
                                        p_2.style = doc.styles['Normal']
                                        for run in p_2.runs:
                                            run.font.size = Pt(11)
                                        break
                                doc_2.save(os.path.abspath(path_ + '\\2 экземпляр' + '\\' + el_))  # Сохраняем
                    percent_val += percent  # Увеличиваем прогресс
                    progress.emit(round(percent_val, 0))  # Посылаем значние в прогресс бар
            if text_for_foot:
                if '/' in num_1:
                    num_1 = text_for_foot.rpartition('/')[0] + '/'
                else:
                    num_1 = text_for_foot.partition('-')[0] + '-'
                try:
                    # Добавляем номер для описи
                    if '/' in num_1:
                        num_2 = str(int((text_for_foot.rpartition('/')[2]).rpartition('c')[0]) + 1)
                    else:
                        num_2 = str(int((text_for_foot.partition('-')[2]).rpartition('c')[0]) + 1)
                except ValueError:
                    if '/' in num_1:
                        num_2 = str(int((text_for_foot.rpartition('/')[2]).rpartition('с')[0]) + 1)
                    else:
                        num_2 = str(int((text_for_foot.partition('-')[2]).rpartition('c')[0]) + 1)
            flag = 0

            # Форма 3
            if flag_inventory == 40:
                if 'Форма 3.docx' in [os.path.basename(i_) for i_ in docs_not]:
                    logging.info("Формируем форму 3")
                    doc = docx.Document(os.path.abspath(path_old + '\\' + 'Форма 3.docx'))  # Открываем
                    text_for_foot = num_1 + num_2 + 'c'  # Текст для нижнего колонитула
                    logging.info("Вставляем колонтитул")
                    insert_header(doc, 11, text_first_header, text_for_foot, hdd_number, executor,
                                  print_people, date, path_, 'Форма 3.docx')
                    if '/' in num_1:  # Добавляем номер для описи
                        num_1 = text_for_foot.rpartition('/')[0] + '/'
                        num_2 = str(int((text_for_foot.rpartition('/')[2]).rpartition('c')[0]) + 1)
                    else:
                        num_1 = text_for_foot.partition('-')[0] + '-'
                        num_2 = str(int((text_for_foot.partition('-')[2]).rpartition('c')[0]) + 1)
                    logging.info("Количество страниц")
                    num_pages = pages_count('Форма 3.docx', path_old)
                    if firm:
                        for_27.append([text_for_foot, date, classified, firm, 'Форма 3.docx', executor,
                                       '1', '№1', str(num_pages - 1)])

            # Если необходимо печатать опись
            if account:  # Если активирована опись
                logging.info("Формируем опись")

                def sort(len_, d):  # Ф-я для записи в необходимом порядке
                    if len_ <= 40:  # Если одной описи хватает для записи документов
                        d.append(len_)  # Добавляем длину последних
                        return d  # Возвращаем
                    else:  # Если не хватает
                        d.append(40)  # Добавляем 40 штук
                        sort(len_ - 40, d)  # Рекурсия

                inventory = 1  # Если выбрана опись
                if flag_inventory == 40:
                    buff = []
                    for el in dict_40:
                        for i_ in el:
                            if re.findall('Заключение', i_) or re.findall('Предписание', i_):
                                buff.append(el)
                    dict_40 = buff
                    len_dict = int(len(dict_40) / 2)  # Получаем длину для записи в опись
                    dict_for_op = []  # Список
                    sort(len_dict, dict_for_op)  # Устанавливаем порядок
                    dict_after = []  # Для записи
                    start_ = 0  # Счетчик
                    for el in dict_for_op:  # Для элементов
                        if flag_inventory == 40:  # Если по 40 в одной описи
                            for i_ in range(0, el):  # Заключения
                                dict_after.append(dict_40[start_ + i_])
                            for i_ in range(0, el):  # Предписания
                                dict_after.append(dict_40[start_ + len_dict + i_])
                        start_ += el
                else:
                    dict_after = dict_40  # Если все в одной описи
                flag_for_op = 0
                percent = 10 / len(dict_after)
                logging.info(dict_after)
                for el in dict_after:  # Для получившихся элементов
                    value = el.popitem()  # Забираем элемент
                    status.emit('Добавляем документ ' + str(value[0]) + ' в опись')
                    name_count = '\\Опись №' + str(inventory) + '.docx'
                    if flag_for_op == 0:  # Если элемент первый в данной описи
                        text_for_foot = num_1 + num_2 + 'c'
                        document = docx.Document()  # Открываем
                        style = document.styles['Normal']
                        font = style.font
                        font.name = 'TimesNewRoman'
                        font.size = Pt(12)
                        section = document.sections[0]
                        # section.orientation, section.page_width, section.page_height
                        new_width, new_height = section.page_height, section.page_width  # Новые размеры
                        section.orientation = WD_ORIENTATION.LANDSCAPE  # Альбомная ориентация
                        section.page_width = new_width
                        section.page_height = new_height
                        section.different_first_page_header_footer = True
                        # Добавляем необходимые надписи перед таблицей, выравниваем, создаем таблицу
                        p = document.add_paragraph('Опись документов № ' + str(inventory))
                        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                        table = document.add_table(rows=1, cols=5, style='Table Grid')
                        style = document.styles['Normal']
                        font = style.font
                        font.name = 'TimesNewRoman'
                        font.size = Pt(12)
                        cell_write(document.styles['Normal'], ['Порядковый номер', 'Наименование документа',
                                                               'Регистрационный номер',
                                                               'Номер экземпляра, гриф секретности',
                                                               'Количество листов в экземпляре'])
                        # Текст внизу таблицы
                        p = document.add_paragraph()
                        p.text = '\n\n' + account_post + '\t\t\t\t\t\t\t\t' + account_signature
                        p.paragraph_format.widow_control = True  # Чтобы подпись не убегала одна
                        p.paragraph_format.keep_together = True  # Чтобы подпись не убегала одна
                        logging.info("Вставляем колонтитул")
                        insert_header(document, 11, value[1][0] + '\n(без приложения не секретно)\nЭкз.№ 1',
                                      text_for_foot, hdd_number, executor,
                                      print_people, date, account_path, name_count, fso)
                        flag_for_op = 1  # Чтобы не создавать, если это не необходимо
                    # Открываем необхоимую опись
                    document = docx.Document(os.path.abspath(account_path + name_count))
                    table = document.tables[0]  # Выбираем таблицу
                    table.add_row()  # Добавляем колонку и значения
                    style = document.styles['Normal']
                    font = style.font
                    font.name = 'TimesNewRoman'
                    font.size = Pt(12)
                    cell_write(document.styles['Normal'], [str(flag_for_op), value[0][:-5] + '...', value[1][1],
                                                           '№' + value[1][2] + ', ' + value[1][0], value[1][3]],
                               flag_for_op)
                    flag_for_op += 1  # Увеличиваем счетчик

                    document.save(account_path + name_count)  # Сохраняем докуент
                    if flag_inventory == 40:  # Если в описи по 40 штук
                        if flag_for_op == 81:  # Если добавили все документы
                            flag_for_op = 0  # Для создания новой описи
                            num_2 = str(int(num_2) + 1)  # Увеличиваем значение для учетного номера
                            inventory += 1  # Номер описи
                    percent_val += percent  # Увеличиваем прогресс
                    flag += 1  # Для того, что бы не мелькал прогресс бар
                    if flag == 4:  # Только для каждого 4 документа при добавлении
                        self.progress.emit(round(percent_val, 0))  # Обновляем прогресс бар
                        flag = 0
                for el in os.listdir(account_path):
                    status.emit('Считаем кол-во листов')
                    if re.findall(r'Опись', el):
                        logging.info("Считаем листы " + el)
                        doc = docx.Document(account_path + '\\' + el)  # Открываем
                        number = doc.sections[0].first_page_footer.paragraphs[0].text
                        num_pages = pages_count(el, account_path)
                        if firm:
                            for_27.append([number, date, classified, firm, el, executor, '1', '№1',
                                           str(num_pages - 1)])

            # Добавление сопровода
            if accompanying_doc:
                logging.info("Добавляем сопровод")
                if not account:
                    text_for_foot = num_1 + num_2 + 'c'
                else:
                    text_for_foot = num_1 + str(int(num_2) + 1) + 'c'
                status.emit('Добавление данных в сопроводительный лист')  # Сообщение в статус бар
                doc = docx.Document(accompanying_doc)
                accompanying_doc = os.path.abspath(path_ + '\\' + os.path.basename(accompanying_doc))
                style = doc.styles['Normal']
                font = style.font
                font.name = 'TimesNewRoman'
                font.size = Pt(11)
                para = True  # Для вставки если сделали не особую первую страницу
                if doc.sections[0].different_first_page_header_footer:
                    header = doc.sections[0].first_page_header  # Верхний колонтитул первой страницы
                    doc.sections[0].footer.paragraphs[0].text = text_for_foot
                else:
                    para = False
                    header = doc.sections[0].header
                head = header.paragraphs[0]  # Параграф
                head.insert_paragraph_before(text_first_header)  # Вставялем перед колонитулом
                head = header.paragraphs[0]  # Выбираем новый первый параграф
                head_format = head.paragraph_format  # Настройки параграфа
                head_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT  # Выравниваем по правому краю
                for p in doc.paragraphs:  # Для каждого параграфа
                    if re.findall(r'registration_number', p.text):  # Ищем метку
                        p.text = re.sub(r'registration_number', text_for_foot + ' от ' + date, p.text)
                        for run in p.runs:
                            run.font.size = Pt(14)
                    elif re.findall(r'Приложения:', p.text):
                        if account:
                            file_account = [i_ for i_ in os.listdir(account_path) if re.findall(r'Опись', i_)]
                            numbering = 1
                            for file in file_account:
                                number = re.findall(r'№(\d*)', file)[0]  # Номер описи
                                docx2pdf.convert(account_path + '\\' + file,
                                                 account_path + '\\' + file + '.pdf')  # Конвертируем
                                input_file = fitz.open(account_path + '\\' + file + '.pdf')  # Открываем пдф
                                pages = input_file.page_count - 1  # Получаем кол-во страниц
                                input_file.close()  # Закрываем
                                os.remove(account_path + '\\' + file + '.pdf')  # Удаляем pdf документ
                                page = 'листе' if pages == 1 else 'листах'  # Для правильной формулировки
                                doc_old = docx.Document(account_path + '\\' + file)  # Открываем
                                footer = doc_old.sections[0].first_page_footer  # Нижний колонтитул первой страницы
                                foot = footer.paragraphs[0]  # Параграф
                                foot_text = foot.text  # Текст нижнего колонитула
                                text = 'Приложение согласно описи №' + str(number) + ' на ' + str(pages) + ' ' \
                                       + page + ', уч. № ' + foot_text + ', экз. № 1, секретно, только в адрес.'
                                p.add_run('\n' + str(numbering) + '. ' + text)
                                p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT  # Выравниваем по левому краю
                                numbering += 1
                                for run in p.runs:
                                    run.font.size = Pt(14)
                                # p.style = doc.styles['Normal']
                        else:
                            file_account = [i_ for i_ in os.listdir(path_) if i_.endswith('.docx') and
                                            ('приложение' not in i_.lower())]
                            if service:
                                file_account = [i_ for i_ in file_account if 'протокол' not in i_.lower()]
                            file_appendix = [i_ for i_ in os.listdir(path_) if 'приложение' in i_.lower()]
                            len_appendix = 0
                            for file in file_appendix:
                                with zipfile.ZipFile(path_ + '\\' + file) as my_doc:
                                    xml_content = my_doc.read('docProps/app.xml')  # Общие свойства
                                    pages = re.findall(r'<Pages>(\w*)</Pages>',
                                                       xml_content.decode())[0]  # Ищем кол-во страниц
                                    len_appendix += int(pages)
                            numbering = 1
                            for file in file_account:
                                with zipfile.ZipFile(path_ + '\\' + file) as my_doc:
                                    xml_content = my_doc.read('docProps/app.xml')  # Общие свойства
                                    pages = re.findall(r'<Pages>(\w*)</Pages>',
                                                       xml_content.decode())[0]  # Ищем кол-во страниц
                                page = 'листе' if int(pages) == 1 else 'листах'
                                doc_old = docx.Document(path_ + '\\' + file)  # Открываем
                                footer = doc_old.sections[0].first_page_footer  # Нижний колонтитул первой страницы
                                foot = footer.paragraphs[0]  # Параграф
                                foot_text = foot.text  # Текст нижнего колонитула
                                text = file.partition(' ')[0] + ', уч. № ' + foot_text + ', экз. № 1, на ' + \
                                       str(int(pages) - 1) + ' ' + page + ' , секретно, только в адрес.'
                                p.add_run('\n' + str(numbering) + '. ' + text)
                                p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT  # Выравниваем по левому краю
                                numbering += 1
                                for run in p.runs:
                                    run.font.size = Pt(14)
                                # p.style = doc.styles['Normal']
                            if len_appendix:
                                page = 'листе' if int(len_appendix) == 1 else 'листах'
                                text = 'Приложение А, на ' + str(len_appendix) + ' ' + page + ' , несекретно.'
                                p.add_run('\n' + str(numbering) + '. ' + text)
                                p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT  # Выравниваем по левому краю
                                numbering += 1
                                for run in p.runs:
                                    run.font.size = Pt(14)
                doc.add_section()  # Добавляем последнюю страницу
                if para:
                    last = doc.sections[len(doc.sections) - 1].first_page_header  # Колонтитул для последней страницы
                    last.is_linked_to_previous = False  # Отвязываем от предыдущей секции чтобы не повторялись
                    foot = doc.sections[len(doc.sections) - 1].first_page_footer  # Нижний колонтитул
                    foot.is_linked_to_previous = False  # Отвязываем
                else:
                    last = doc.sections[len(doc.sections) - 1].header  # Колонтитул для последней страницы
                    last.is_linked_to_previous = False  # Отвязываем от предыдущей секции чтобы не повторялись
                    foot = doc.sections[len(doc.sections) - 1].footer  # Нижний колонтитул
                    foot.is_linked_to_previous = False  # Отвязываем
                # Текст для фонарика
                foot.paragraphs[0].text = "Уч. № " + text_for_foot + \
                                          "\nОтп. 2 экз.\n№ 1 - в адрес\n№ 2 - в дело \nс НЖМД, уч. № ЖД - " \
                                          + hdd_number + "\nИсп. " \
                                          + executor_acc_sheet + "\nПеч. " + print_people + \
                                          "\n" + date + "\nБ/ч"
                doc.save(accompanying_doc)  # Сохраняем
                num_pages = pages_count(os.path.basename(accompanying_doc),
                                        accompanying_doc.rpartition('\\')[0])
                if firm:
                    for_27.append([text_for_foot, date, classified, firm,
                                   accompanying_doc.rpartition('\\')[2][:-5],
                                   executor_acc_sheet, '1', '№1', str(num_pages - 1)])
            if '/' in num_1:  # Добавляем номер для описи
                num_2 = str(int((text_for_foot.rpartition('/')[2]).rpartition('c')[0]) + 1)  # Увеличиваем номер
            else:
                num_2 = str(int((text_for_foot.partition('-')[2]).rpartition('c')[0]) + 1)
            # Добавляем форму 27
            if firm:
                logging.info("Формируем форму 27")
                percent = 10 / (len(os.listdir()))  # Процент для прогресса
                status.emit('Добавление данных в 27 форму')  # Сообщение в статус бар
                i_ = 0
                table_name = ['Порядковый номер', 'Дата регистрации',
                              'Номер, дата поступившего документа и гриф секретности',
                              'Откуда (от кого) поступил или кому направлен документ',
                              'Наименование или краткое содержание документа', 'Фамилия исполнителя и подразделение',
                              'экземпляров и их номера', 'листов в экземпляре', 'листов основного документа',
                              'листов приложения',
                              'Номера блока и листов черновика', 'Номера перепечатанных листов',
                              'Отметка об уничтожении черновиков', 'Кому выдан документ',
                              'Расписка в получении документа и дата', 'Номер реестра и дата',
                              'Местонахождение документа(номер дела и листа, номер акта на уничтожение и дата)',
                              'Примечание']
                table_df = pd.DataFrame({'Порядковый номер': [],
                                         'Дата регистрации': [],
                                         'Номер, дата поступившего документа и гриф секретности': [],
                                         'Откуда (от кого) поступил или кому направлен документ': [],
                                         'Наименование или краткое содержание документа': [],
                                         'Фамилия исполнителя и подразделение': [],
                                         'экземпляров и их номера': [],
                                         'листов в экземпляре': [],
                                         'листов основного документа': [],
                                         'листов приложения': [],
                                         'Номера блока и листов черновика': [],
                                         'Номера перепечатанных листов': [],
                                         'Отметка об уничтожении черновиков': [],
                                         'Кому выдан документ': [],
                                         'Расписка в получении документа и дата': [],
                                         'Номер реестра и дата': [],
                                         'Местонахождение документа(номер дела и листа, номер акта на уничтожение'
                                         ' и дата)': [],
                                         'Примечание': []})
                for element in for_27:
                    status.emit('Добавление данных в 27 форму (' + element[4] + ')')  # Сообщение в статус бар
                    num = 0  # Номер столбца
                    if element[6] == '2':
                        table_df.loc[(i_ - 3), table_name[6]] = '2'
                        table_df.loc[i_, table_name[6]] = '№2'
                        table_df.loc[i_, table_name[7]] = element[8]
                        i_ += 1
                        table_df.loc[i_] = pd.Series([numpy.NaN for i_ in range(0, len(table_name))], index=table_name)
                        i_ += 1
                    else:
                        for el in element:
                            table_df.loc[i_, table_name[num]] = el
                            num += 1
                            if element.index(el) == 6 and num <= 7:
                                num -= 1
                                i_ += 1
                        i_ += 1
                        if re.findall(r'сопроводит', element[4].lower()):
                            table_df.loc[i_] = pd.Series([numpy.NaN for i_ in range(0, len(table_name))], index=table_name)
                            i_ += 1
                            table_df.loc[i_, table_name[6]] = '№2'
                            table_df.loc[i_, table_name[7]] = table_df.loc[(i_ - 2), table_name[7]]
                            i_ += 1
                            table_df.loc[i_] = pd.Series([numpy.NaN for i_ in range(0, len(table_name))], index=table_name)
                        if for_27.index(element) != len(for_27) - 1:
                            table_df.loc[i_] = pd.Series([numpy.NaN for i_ in range(0, len(table_name))], index=table_name)
                            i_ += 1
                    percent_val += percent  # Увеличиваем прогресс
                    progress.emit(round(percent_val, 0))  # Обновляем прогресс бар
                table_df.index = pd.RangeIndex(1, 1 + len(table_df))
                table_df.to_excel(path_ + '\\Форма 27.xlsx', sheet_name='27', index=False)
                column_width = [13, 11, 10, 24, 27, 13.5, 7, 7, 7, 7, 15.1, 13, 13.1, 11.4, 16, 18.3, 23.85,
                                14.1]
                wb_ = openpyxl.load_workbook(path_ + '\\Форма 27.xlsx')
                ws_ = wb_.active
                ws_.insert_rows(2)
                thin = openpyxl.styles.Side(border_style="thin", color="000000")
                for el in range(1, ws_.max_column + 1):
                    if 6 < el < 11:
                        ws_.cell(2, el).value = ws_.cell(1, el).value
                        ws_.cell(2, el).alignment = openpyxl.styles.Alignment(horizontal="center",
                                                                              vertical="center", wrap_text=True)
                        ws_.cell(2, el).border = openpyxl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
                    else:
                        if el == 11:
                            ws_.merge_cells(start_row=1, end_row=1, start_column=7, end_column=10)
                            ws_.cell(1, 7).value = 'Количество'
                            ws_.cell(1, 7).alignment = openpyxl.styles.Alignment(horizontal="center",
                                                                                 vertical="center", wrap_text=True)
                        ws_.merge_cells(start_row=1, end_row=2, start_column=el, end_column=el)
                        ws_.cell(1, el).alignment = openpyxl.styles.Alignment(horizontal="center",
                                                                              vertical="center", wrap_text=True)
                for el in range(1, ws_.max_column + 1):
                    ws_.column_dimensions[get_column_letter(el)].width = column_width[el - 1]
                flag = 0
                for row in range(3, ws_.max_row + 1):
                    flag += 1
                    for col in range(1, ws_.max_column + 1):
                        if flag == 4:
                            flag = 1
                        if flag == 1:
                            if col == 3 and ws_.cell(row, col).value:
                                ws_.cell(row, col).alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="top")
                                ws_.merge_cells(start_row=row, end_row=row + 1, start_column=col, end_column=col)
                        if ws_.cell(row, col).value == 0:
                            ws_.cell(row, col).value = ''
                        ws_.cell(row, col).border = openpyxl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
                wb_.save(filename=path_form_27 + '\\Форма 27.xlsx') if path_form_27 else wb_.save(
                    filename=path_ + '\\Форма 27.xlsx')

                def form_27_name(path_wb_f27):
                    wb_f27 = openpyxl.load_workbook(path_wb_f27)
                    ws_f27 = wb_f27.active
                    number_f27 = [ws_f27.cell(i_f27, 1).value for i_f27 in range(2, ws_f27.max_row + 1) if
                                  ws_f27.cell(i_f27, 1).value]
                    number_f27 = natsorted(number_f27)
                    wb_f27.close()
                    name_wb = '\\Форма 27' \
                              ' ' + str(number_f27[0]).replace('/', ',') + \
                              ' - ' + str(number_f27[-1]).replace('/', ',') + '.xlsx'
                    os.rename(path_wb_f27, os.path.dirname(path_wb_f27) + name_wb)

                path_f27 = path_form_27 + '\\Форма 27.xlsx' if path_form_27 else path_ + '\\Форма 27.xlsx'
                form_27_name(path_f27)
            return num_1, num_2
        time_start = datetime.datetime.now()
        self.progress.emit(0)  # Обновляем статус бар
        self.status.emit('Начинаем')
        self.logging.info("Старт программы")
        pt_num = 14 if self.service else 12
        try:  # Для отлова ошибок
            if self.file_num:  # Если есть файл номеров
                dict_file = {}
                wb = openpyxl.load_workbook(self.file_num)  # Откроем книгу.
                ws = wb.active  # Делаем активным первый лист.
                for i in range(1, ws.max_row + 1):  # Пока есть значения
                    if ws.cell(i, 1).value:
                        cv = ws.cell(i, 1).value.partition(' ')[0].lower()  # Получаем значние
                        if cv == 'заключение' or cv == 'предписание' or cv == 'протокол' or cv == 'форма' \
                                or cv == 'опись' or cv == 'сопроводит':
                            dict_file[ws.cell(i, 1).value] = [ws.cell(i, 2).value + 'c',
                                                              ws.cell(i, 3).value.strftime("%d.%m.%Y")]  # Делаем список
            else:  # Если нет файла номеров
                if re.match(r'\w+/\w+/\w+c', self.number):
                    self.num_1 = self.number.rpartition('/')[0] + '/'
                    self.num_2 = self.number.rpartition('/')[2].rpartition('c')[0]
                else:
                    self.num_1 = self.number.partition('-')[0] + '-'
                    self.num_2 = self.number.partition('-')[2].rpartition('c')[0]
            self.logging.info("Созданы секретные номера")
            if self.package:
                for folder in os.listdir(self.path_old):
                    os.chdir(self.path_old + '\\' + folder)
                    path_old = self.path_old + '\\' + folder
                    path = self.path_new + '\\' + folder
                    try:
                        os.mkdir(path)
                    except FileExistsError:
                        self.status.emit('Ошибка')  # Сообщение в статус бар
                        self.messageChanged.emit('УПС!', 'В конечной папке уже присутствует папка «'
                                                 + path.rpartition('\\')[2] + '». Удалите или переместите её.')
                        return
                    self.num_1, self. num_2 = format_doc_(path_old, self.classified, self.list_item, self.num_scroll,
                                                          self.account, self.firm, self.logging, self.status, path,
                                                          self.file_num, self.num_1, self.num_2, self.date,
                                                          self.conclusion, self.executor, self.prescription,
                                                          self.hdd_number, self.print_people, self.progress,
                                                          self.flag_inventory, self.account_post,
                                                          self.account_signature, self.account_path,
                                                          self.executor_acc_sheet, self.service, False)
                    docs_txt = [file for file in os.listdir(path_old) if file[-4:] == '.txt']  # Список txt
                    for txt_file in docs_txt:
                        shutil.copy(txt_file, path)
            else:
                format_doc_(self.path_old, self.classified, self.list_item, self.num_scroll, self.account,
                            self.firm, self.logging, self.status, self.path_new, self.file_num, self.num_1, self.num_2,
                            self.date, self.conclusion, self.executor, self.prescription, self.hdd_number,
                            self.print_people, self.progress, self.flag_inventory,
                            self.account_post, self.account_signature, self.account_path, self.executor_acc_sheet,
                            self.service, self.path_form_27)
                docs_txt = [file for file in os.listdir(self.path_old) if file[-4:] == '.txt']  # Список txt
                for txt_file in docs_txt:
                    shutil.copy(txt_file, self.path_new)
            self.logging.info("Конец программы, время работы: " + str(datetime.datetime.now() - time_start))
            self.logging.info("\n***********************************************************************************\n")
            self.status.emit('Готово!')  # Посылаем значние если готово
            self.progress.emit(100)  # Завершаем прогресс бар
        except BaseException as e:  # Если ошибка
            self.status.emit('Ошибка')  # Сообщение в статус бар
            self.logging.error("Ошибка:\n " + str(e) + '\n' + traceback.format_exc())
