import os
import traceback
import win32api
import win32print
from pathlib import Path


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
                printer_defaults = {"DesiredAccess": win32print.PRINTER_ACCESS_USE}  # Дефолтный принтер
                handle = win32print.OpenPrinter(name_printer, printer_defaults)  # Открываем
                attributes = win32print.GetPrinter(handle, 1)
                attributes['pDevMode'].Duplex = 1  # flip up  Для двухсторонней печати
                try:
                    # Устанавливаем настройки
                    win32print.SetPrinter(handle, 1, attributes, 0)
                except:  # Пропускаем ошибку
                    pass
                win32api.ShellExecute(0, "print",
                                      str(Path(incoming_data['start_path'], file)), name_printer, ".", 0)
                jobs = 0  # Проверка для того, что бы не перескакивать на следующий документ
                logging.info(f"Ждем очередь")
                while jobs < 3:
                    print_jobs = win32print.EnumJobs(handle, 0, -1, 1)  # Очередь печати
                    if not print_jobs and jobs == 0:  # Пока не запустилось в печать
                        pass
                    elif not print_jobs and jobs == 2:  # Если запустилось и очистилась
                        jobs = 3
                        logging.info('Очередь очистилась')
                    elif print_jobs:  # Если в очереди что-то есть
                        jobs = 2
                win32print.ClosePrinter(handle)  # Закрываем принтер
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