import re
import time
import traceback

import docx
import fitz
import os
from pathlib import Path

import pythoncom

from word2pdf import word2pdf


def insert_sign2pdf(incoming_data: dict, current_progress, now_doc, all_doc, line_doing, line_progress, progress_value,
                    event, window_check, info_value) -> dict:
    try:
        errors = []
        logging = incoming_data['logging']
        percent = 100 / all_doc
        logging.info(f"Бежим по папке {Path(incoming_data['start_path']).name}")
        for file in os.listdir(incoming_data['start_path']):
            try:
                event.wait()
                if window_check.stop_threading:
                    return {'status': 'cancel', 'trace': '', 'text': ''}
                if not Path(incoming_data['start_path'], file).is_file():
                    continue
                if file.endswith('.docx') is False:
                    continue
                line_doing.emit(f'Вставляем подпись в файл {file} ({now_doc} из {all_doc})')
                name_pdf = file.rpartition('.')[0] + '.pdf'
                pdf_path = Path(incoming_data['start_path'], name_pdf)
                word_path = Path(incoming_data['start_path'], file)
                finish_pdf_path = Path(incoming_data['finish_path'], name_pdf)
                pythoncom.CoInitializeEx(0)
                document = docx.Document(str(word_path))
                pattern = re.compile(r'\{\{\s[A-z]*\s}}')
                find_name = [pattern.findall(p.text)[0] for p in document.paragraphs if pattern.findall(p.text)]
                if len(find_name):
                    word2pdf(str(word_path), str(pdf_path))
                    doc = fitz.open(str(pdf_path))
                    insert_text = {}
                    for index in range(len(doc)):
                        page = doc.load_page(index)
                        text_instances = {name[3: len(name) - 3].lower(): page.search_for(name) for name in find_name}
                        if len(text_instances) > 0:
                            insert_text[index] = text_instances
                    doc.close()
                    # os.remove(pdf_path)
                    # word2pdf(str(word_path), str(pdf_path))
                    doc = fitz.open(str(pdf_path))
                    for index in insert_text:
                        page = doc.load_page(index)
                        for inst in insert_text[index]:
                            for i in insert_text[index][inst]:
                                if inst == 'seal':
                                    rect = fitz.Rect(i.x0 - 60, i.y0 - 60, i.x1 + 60, i.y1 + 60)
                                else:
                                    rect = fitz.Rect(i.x0 - 32, i.y0 - 32, i.x1 + 32, i.y1 + 32)
                                page.insert_image(rect, filename=str(incoming_data['signature_files'][inst]))
                    doc.save(str(finish_pdf_path))
                    doc.close()
                    permission = 0
                    while True:
                        try:
                            os.remove(pdf_path)
                            break
                        except PermissionError as err:
                            logging.warning(f"Ошибка при удалении {pdf_path.name} - {err}\n{traceback.format_exc()}")
                            time.sleep(3)
                            doc = fitz.open(str(pdf_path))
                            doc.close()
                        if permission > 3:
                            break
                        permission += 1
                    if permission == 4:
                        logging.warning(f"Не удалось удалить pdf файл {pdf_path.name}")
                current_progress += percent
                line_progress.emit(f'Выполнено {int(current_progress)} %')
                progress_value.emit(int(current_progress))
                now_doc += 1
            except BaseException as error:
                errors.append(f"Ошибка при вставке подписи в {file}")
                logging.warning(f"Ошибка при вставке подписи {file} - {error}\n{traceback.format_exc()}")
        line_progress.emit(f'Выполнено {int(100)} %')
        progress_value.emit(int(100))
        return {'status': 'warning' if errors else 'success', 'text': errors, 'trace': ''}
    except BaseException as error:
        return {'status': 'error', 'text': error, 'trace': traceback.format_exc()}