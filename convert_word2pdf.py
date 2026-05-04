import traceback
import pythoncom
from pathlib import Path

import win32com

from small_functions import replace_object
from word2pdf import word2pdf


def convert_word2pdf(incoming_data: dict, current_progress, now_doc, all_doc, line_doing, line_progress,
                     progress_value, event, window_check, info_value) -> dict:
    try:
        logging = incoming_data['logging']
        errors = []
        percent = 100/all_doc
        for file in Path(incoming_data['start_path']).rglob('*.*'):
            if file.suffix != '.docx':
                continue
            event.wait()
            if window_check.stop_threading:
                return {'status': 'cancel', 'trace': '', 'text': ''}
            pythoncom.CoInitializeEx(0)
            line_doing.emit(f'Преобразуем Word {file.name} в PDF ({now_doc} из {all_doc})')
            logging.info(f'Преобразуем {file.name} в PDF')
            convert_file = Path(incoming_data['finish_path'], file.stem + '.pdf')
            replace = replace_object(convert_file, logging, info_value, event, window_check)
            if replace:
                try:
                    word2pdf(str(file), str(convert_file))
                except BaseException:
                    word = win32com.client.Dispatch("Word.Application")
                    word.Quit()
                    try:
                        word2pdf(str(file), str(convert_file))
                    except BaseException as ex:
                        logging.error(f"Ошибка при преобразовании {file.name} в PDF: {ex}")
                        errors.append(f"Ошибка при преобразовании {file.name}, проверьте наличие PDF")
            current_progress += percent
            line_progress.emit(f'Выполнено {int(current_progress)} %')
            progress_value.emit(int(current_progress))
            now_doc += 1
        return {'status': 'warning' if errors else 'success', 'text': errors, 'trace': ''}
    except BaseException as error:
        return {'status': 'error', 'text': error, 'trace': traceback.format_exc()}
