import os
import traceback
from pathlib import Path
from small_functions import print_doc


def print_all_files(incoming_data: dict, current_progress, now_doc, all_doc, line_doing, line_progress, progress_value,
                    event, window_check, info_value) -> dict:
    """Печать всех файлов из указанной папки."""
    try:
        errors = []
        logging = incoming_data['logging']
        name_printer = incoming_data['printer']
        percent = 100 / all_doc
        logging.info(f"Бежим по папке {Path(incoming_data['start_path']).name}")
        for file in os.listdir(incoming_data['start_path']):
            try:
                event.wait()
                if window_check.stop_threading:
                    return {'status': 'cancel', 'trace': '', 'text': ''}
                if not Path(incoming_data['start_path'], file).is_file():
                    current_progress += percent
                    line_progress.emit(f'Выполнено {int(current_progress)} %')
                    progress_value.emit(int(current_progress))
                    now_doc += 1
                    continue
                line_doing.emit(f'Печатаем {file} ({now_doc} из {all_doc})')
                answer = print_doc(Path(incoming_data['start_path'], file), name_printer, 1, logging)
                if answer['status'] != 'success':
                    errors.append(answer['text'])
                    logging.error(answer['text'])
                if answer['status'] == 'error':
                    logging.error(answer['trace'])
                current_progress += percent
                line_progress.emit(f'Выполнено {int(current_progress)} %')
                progress_value.emit(int(current_progress))
                now_doc += 1
            except BaseException as error:
                errors.append(f"Ошибка при печати {file}, файл не напечатан")
                logging.warning(f"Ошибка при печати файла {file} - {error}\n{traceback.format_exc()}")
        line_progress.emit(f'Выполнено {int(100)} %')
        progress_value.emit(int(100))
        return {'status': 'warning' if errors else 'success', 'text': errors, 'trace': ''}
    except BaseException as error:
        return {'status': 'error', 'text': error, 'trace': traceback.format_exc()}
