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


def print_certification(incoming_data: dict, current_progress, now_doc, all_doc, line_doing,
                        line_progress, progress_value, event, window_check, info_value) -> dict:
    """Печать для лаборатории сертификации."""
    try:
        errors = []
        logging = incoming_data['logging']
        name_printer = incoming_data['name_printer']
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
        logging.info(f"Бежим по папке {Path(incoming_data['start_path']).name}")
        for file in os.listdir(incoming_data['start_path']):
            try:
                event.wait()
                if window_check.stop_threading:
                    return {'status': 'cancel', 'trace': '', 'text': ''}
                file_path = Path(incoming_data['start_path'], file)
                pdf_path = Path(incoming_data['start_path'], file_path.stem + '.pdf')
                if not file_path.is_file():
                    current_progress += percent
                    line_progress.emit(f'Выполнено {int(current_progress)} %')
                    progress_value.emit(int(current_progress))
                    now_doc += 1
                    continue
                line_doing.emit(f'Печатаем {file} ({now_doc} из {all_doc})')
                printing_date = [computer_name, user_name, str(file_path), str(datetime.date.today()), printer]
                pythoncom.CoInitializeEx(0)
                if re.findall(r'приложение а', file.lower()):
                    logging.info(f"Запускаем в печать {file}")
                    answer = print_doc(file_path, name_printer, 1, logging)
                    if answer['status'] != 'success':
                        errors.append(answer['text'])
                        logging.error(answer['text'])
                    if answer['status'] == 'error':
                        logging.error(answer['trace'])
                else:
                    try:
                        word2pdf(str(file_path), str(pdf_path))
                    except BaseException:
                        word = win32com.client.Dispatch("Word.Application")
                        word.Quit()
                    if incoming_data['one_side']:  # Если печать односторонняя - печатаем
                        logging.info(f"Печатаем документ {file}")
                        answer = print_doc(pdf_path, name_printer, 1, logging)
                        if answer['status'] != 'success':
                            errors.append(answer['text'])
                            logging.error(answer['text'])
                        if answer['status'] == 'error':
                            logging.error(answer['trace'])
                    elif incoming_data['last_duplex']:  # Если последняя страница дуплекс
                        # Дефолтный принтер
                        logging.info(f"Преобразуем документ {file}")
                        input_file = fitz.open(pdf_path)  # Открываем пдф
                        pages = input_file.page_count  # Получаем кол-во страниц
                        if pages < 2:
                            logging.warning(f"Выбрана печать двухсторонним методом, но в файле {file} {pages} страниц."
                                            f" Печатаем как есть")
                            answer = print_doc(pdf_path, name_printer, 1, logging)
                            if answer['status'] != 'success':
                                errors.append(answer['text'])
                                logging.error(answer['text'])
                            if answer['status'] == 'error':
                                logging.error(answer['trace'])
                        elif pages == 2:
                            answer = print_doc(pdf_path, name_printer, 2, logging)
                            if answer['status'] != 'success':
                                errors.append(answer['text'])
                                logging.error(answer['text'])
                            if answer['status'] == 'error':
                                logging.error(answer['trace'])
                        else:
                            logging.info(f"Преобразуем документ {pdf_path}")
                            output_1_side = Path(pdf_path.parent, '1_' + pdf_path.name)
                            output_2_side = Path(pdf_path.parent, '2_' + pdf_path.name)
                            # Страницы для односторонней печати
                            selected_page = [page for page in range(0, pages - 2)]
                            input_file.select(selected_page)  # Выбираем страницы
                            input_file.save(output_1_side)  # Сохраняем файл
                            # Печатаем
                            answer = print_doc(output_1_side, name_printer, 1, logging)
                            if answer['status'] != 'success':
                                errors.append(answer['text'])
                                logging.error(answer['text'])
                            if answer['status'] == 'error':
                                logging.error(answer['trace'])
                            os.remove(output_1_side)
                            input_file = fitz.open(pdf_path)  # Открываем пдф
                            selected_page = [pages - 2, pages - 1]  # Страницы для двухсторонней печати
                            input_file.select(selected_page)  # Выбираем страницы
                            input_file.save(output_2_side)  # Сохраняем
                            answer = print_doc(output_2_side, name_printer, 2, logging)
                            if answer['status'] != 'success':
                                errors.append(answer['text'])
                                logging.error(answer['text'])
                            if answer['status'] == 'error':
                                logging.error(answer['trace'])
                            os.remove(output_2_side)
                        input_file.close()
                    else:
                        answer = print_doc(pdf_path, name_printer, 2, logging)
                        if answer['status'] != 'success':
                            errors.append(answer['text'])
                            logging.error(answer['text'])
                        if answer['status'] == 'error':
                            logging.error(answer['trace'])
                    logging.info('Записываем данные с печати')
                    with open(save_printing_data_file, 'a') as f:
                        f.write(';'.join(printing_date) + '\n')
            except Exception as ex:
                logging.error("Упс, сорвалась печать файла!")
                logging.error("Ошибка:\n " + str(ex) + '\n' + traceback.format_exc())
            pdf_files = [i for i in os.listdir(incoming_data['start_path']) if file.rpartition('.')[0] in i
                         and i.endswith('.pdf')]
            if pdf_files:
                logging.info(f"Удаляем пдф после печати")
                for pdf_file in pdf_files:
                    permission = 0
                    while permission < 4:
                        try:
                            permission += 1
                            os.remove(str(Path(incoming_data['start_path'], pdf_file)))
                            break
                        except PermissionError as per:
                            time.sleep(3)
                            logging.warning(f"Ошибка удаления файла {pdf_file} после печати - {per}")
                            input_file = fitz.open(Path(incoming_data['start_path'], pdf_file))  # Открываем пдф
                            input_file.close()
            current_progress += percent
            line_progress.emit(f'Выполнено {int(current_progress)} %')
            progress_value.emit(int(current_progress))
            now_doc += 1
        line_progress.emit(f'Выполнено {int(100)} %')
        progress_value.emit(int(100))
        return {'status': 'warning' if errors else 'success', 'text': errors, 'trace': ''}
    except BaseException as error:
        return {'status': 'error', 'text': error, 'trace': traceback.format_exc()}
