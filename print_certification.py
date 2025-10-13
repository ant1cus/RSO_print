import datetime
import getpass
import os
import re
import socket
import time
import traceback
import pythoncom
from pathlib import Path

import fitz
import win32com

from small_functions import print_doc
from word2pdf import word2pdf

from PyQt5 import QtPrintSupport


def print_file(start_path: Path, file: str, computer_name: str, user_name: str, printer: str,
               save_printing_data_file: Path, incoming_data: dict, now_doc, all_doc, line_doing,
               event, window_check) -> dict:
    file_path = Path(start_path, file)
    pdf_path = Path(start_path, file_path.stem + '.pdf')
    error = ''
    try:
        event.wait()
        if window_check.stop_threading:
            return {'status': 'cancel', 'trace': '', 'text': ''}
        if not file_path.is_file():
            return {'status': 'ok', 'trace': '', 'text': ''}
        line_doing.emit(f'Печатаем {file} ({now_doc} из {all_doc})')
        printing_date = [computer_name, user_name, str(file_path), str(datetime.date.today()), printer]
        pythoncom.CoInitializeEx(0)
        if re.findall(r'приложение а', file.lower()):
            incoming_data['logging'].info(f"Запускаем в печать {file}")
            answer = print_doc(file_path, incoming_data['name_printer'], 1, incoming_data['logging'])
            if answer['status'] != 'success':
                error = answer['text']
                incoming_data['logging'].error(answer['text'])
            if answer['status'] == 'error':
                incoming_data['logging'].error(answer['trace'])
        else:
            try:
                word2pdf(str(file_path), str(pdf_path))
            except BaseException:
                word = win32com.client.Dispatch("Word.Application")
                word.Quit()
            if incoming_data['one_side']:  # Если печать односторонняя - печатаем
                incoming_data['logging'].info(f"Печатаем документ {file}")
                answer = print_doc(pdf_path, incoming_data['name_printer'], 1, incoming_data['logging'])
                if answer['status'] != 'success':
                    error = answer['text']
                    incoming_data['logging'].error(answer['text'])
                if answer['status'] == 'error':
                    incoming_data['logging'].error(answer['trace'])
            elif incoming_data['last_duplex']:  # Если последняя страница дуплекс
                # Дефолтный принтер
                incoming_data['logging'].info(f"Преобразуем документ {file}")
                input_file = fitz.open(pdf_path)  # Открываем пдф
                pages = input_file.page_count  # Получаем кол-во страниц
                if pages < 2:
                    incoming_data['logging'].warning(f"Выбрана печать двухсторонним методом,"
                                                     f" но в файле {file} {pages} страниц."
                                                     f" Печатаем как есть")
                    answer = print_doc(pdf_path, incoming_data['name_printer'], 1, incoming_data['logging'])
                    if answer['status'] != 'success':
                        error = answer['text']
                        incoming_data['logging'].error(answer['text'])
                    if answer['status'] == 'error':
                        incoming_data['logging'].error(answer['trace'])
                elif pages == 2:
                    answer = print_doc(pdf_path, incoming_data['name_printer'], 2, incoming_data['logging'])
                    if answer['status'] != 'success':
                        error = answer['text']
                        incoming_data['logging'].error(answer['text'])
                    if answer['status'] == 'error':
                        incoming_data['logging'].error(answer['trace'])
                else:
                    incoming_data['logging'].info(f"Преобразуем документ {pdf_path}")
                    output_1_side = Path(pdf_path.parent, '1_' + pdf_path.name)
                    output_2_side = Path(pdf_path.parent, '2_' + pdf_path.name)
                    # Страницы для односторонней печати
                    selected_page = [page for page in range(0, pages - 2)]
                    input_file.select(selected_page)  # Выбираем страницы
                    input_file.save(output_1_side)  # Сохраняем файл
                    # Печатаем
                    answer = print_doc(output_1_side, incoming_data['name_printer'], 1, incoming_data['logging'])
                    if answer['status'] != 'success':
                        error = answer['text']
                        incoming_data['logging'].error(answer['text'])
                    if answer['status'] == 'error':
                        incoming_data['logging'].error(answer['trace'])
                    os.remove(output_1_side)
                    input_file = fitz.open(pdf_path)  # Открываем пдф
                    selected_page = [pages - 2, pages - 1]  # Страницы для двухсторонней печати
                    input_file.select(selected_page)  # Выбираем страницы
                    input_file.save(output_2_side)  # Сохраняем
                    answer = print_doc(output_2_side, incoming_data['name_printer'], 2, incoming_data['logging'])
                    if answer['status'] != 'success':
                        error = answer['text']
                        incoming_data['logging'].error(answer['text'])
                    if answer['status'] == 'error':
                        incoming_data['logging'].error(answer['trace'])
                    os.remove(output_2_side)
                input_file.close()
            else:
                answer = print_doc(pdf_path, incoming_data['name_printer'], 2, incoming_data['logging'])
                if answer['status'] != 'success':
                    error = answer['text']
                    incoming_data['logging'].error(answer['text'])
                if answer['status'] == 'error':
                    incoming_data['logging'].error(answer['trace'])
            incoming_data['logging'].info('Записываем данные с печати')
            with open(save_printing_data_file, 'a') as f:
                f.write(';'.join(printing_date) + '\n')
    except Exception as ex:
        incoming_data['logging'].error("Упс, сорвалась печать файла!")
        incoming_data['logging'].error("Ошибка:\n " + str(ex) + '\n' + traceback.format_exc())
    try:
        incoming_data['logging'].info(f"Удаляем пдф после печати")
        permission = 0
        while permission < 4:
            try:
                permission += 1
                os.remove(str(pdf_path))
                break
            except PermissionError as per:
                time.sleep(3)
                incoming_data['logging'].warning(f"Ошибка удаления файла {pdf_path.name} после печати - {per}")
                input_file = fitz.open(pdf_path)  # Открываем пдф
                input_file.close()
        if permission == 4:
            error = f"Не удалось удалить pdf {pdf_path.name}"
    except Exception as ex:
        incoming_data['logging'].error("Упс, не удалось удалить пдф после печати!")
        incoming_data['logging'].error("Ошибка:\n " + str(ex) + '\n' + traceback.format_exc())
    return {'status': 'ok', 'error': error, 'trace': '', 'text': ''}


def print_certification(incoming_data: dict, current_progress, now_doc, all_doc, line_doing,
                        line_progress, progress_value, event, window_check, info_value) -> dict:
    """Печать для лаборатории сертификации."""
    try:
        errors = []
        logging = incoming_data['logging']
        user_name = getpass.getuser()
        printer = QtPrintSupport.QPrinterInfo.defaultPrinterName()
        computer_name = socket.gethostname()
        date_for_saving = datetime.date.today()
        if not Path.exists(Path(incoming_data['default_path'], 'printing_data')):
            os.makedirs(Path(incoming_data['default_path'], 'printing_data'))
        save_printing_data_file = Path(incoming_data['default_path'], 'printing_data', str(date_for_saving) + '.txt')
        if not os.path.exists(save_printing_data_file):
            with open(save_printing_data_file, 'w'):
                pass
        percent = 100 / all_doc
        if Path(incoming_data['start_path']).is_file():
            print_file(Path(incoming_data['start_path']).parent, incoming_data['start_path'], computer_name,
                       user_name, printer, save_printing_data_file, incoming_data, now_doc, all_doc, line_doing,
                       event, window_check)
        else:
            logging.info(f"Бежим по папке {Path(incoming_data['start_path']).name}")
            for file in os.listdir(incoming_data['start_path']):
                answer = print_file(Path(incoming_data['start_path']), file, computer_name, user_name, printer,
                                    save_printing_data_file, incoming_data, now_doc, all_doc, line_doing,
                                    event, window_check)
                if answer['status'] == 'cancel':
                    return {'status': 'cancel', 'trace': '', 'text': ''}
                if len(answer['error']) > 0:
                    errors.append(answer['error'])
                current_progress += percent
                line_progress.emit(f'Выполнено {int(current_progress)} %')
                progress_value.emit(int(current_progress))
                now_doc += 1
        line_progress.emit(f'Выполнено {int(100)} %')
        progress_value.emit(int(100))
        return {'status': 'warning' if errors else 'success', 'text': errors, 'trace': ''}
    except BaseException as error:
        return {'status': 'error', 'text': error, 'trace': traceback.format_exc()}
