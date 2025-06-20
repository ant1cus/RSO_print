import os
import re
import shutil
import traceback
from pathlib import Path

from get_text_num import create_text_for_docs
from small_functions import sorting_files, pages_count, inventory_insert, sp_sorting, sort_order
from create_form_27 import create_form_27
from create_files import create_file

import pandas as pd


def add_documents(incoming_data: dict, start_path: Path, finish_path: Path, line_doing, line_progress, progress_value,
                  event, window_check, info_value) -> dict:
    """Для каждой отдельной папки"""
    try:
        # start_path - начальный путь (каждая папка), finish_path - конечный путь (для каждой папки свой)
        # columns_name [имя, номер, где находится, куда положить, родительская папка,
        # является ли документ приложением а,
        # что делать с документом (копирование, и что-то ещё забыл, надо подумать надо ли),
        # входит ли в опись, текст в верхних полях, надо ли менять дату в документе (надо сделать просто везде сразу),
        # исполнитель, текст для верхнего колонтитула, текст для нижнего колонтитула, дата,
        # текст для вставки в заключение, текст для вставки в предписание, фонарик, кол-во страниц]
        columns_name = ['name', 'number', 'start_path', 'finish_path', 'parent_path', 'a_prescription', 'action',
                        'account_list', 'text', 'change_date', 'executor',
                        'first_header_text', 'footer_text', 'date', 'text_conclusion', 'text_protocol',
                        'text_finish', 'pages', 'classified', 'num_scroll', 'account_list_text', 'form_27']
        documents = pd.DataFrame(columns=columns_name)
        errors = []
        logging = incoming_data['logging']
        for file in Path(start_path).rglob('*.*'):
            documents = pd.concat([documents, pd.DataFrame({'name': [file.name], 'start_path': [file],
                                                            'finish_path': [Path(finish_path, file.name)],
                                                            'parent_path': [file.parent],
                                                            'classified': [incoming_data['classified']],
                                                            'num_scroll': [incoming_data['num_scroll']],
                                                            'date': [incoming_data['date']]})], ignore_index=True)
        inventory = []
        if incoming_data['inventory_insert']:
            logging.info("Добавляем опись")
            number_inventory = 1
            inventory = []
            if incoming_data['flag_inventory'] == 1:
                inventory.append(Path(incoming_data['start_path'], f"Опись №{number_inventory}.docx"))
                documents = pd.concat([documents, pd.DataFrame({'name': [f"Опись №{number_inventory}.docx"]})],
                                      ignore_index=True)
            else:
                conclusion_num = len(documents[(documents['name'].str.contains(r'заключение', case=False))])
                while True:
                    if conclusion_num <= 40:  # Если одной описи хватает для записи документов
                        inventory.append(Path(incoming_data['start_path'], f"Опись №{number_inventory}.docx"))
                        documents = pd.concat([documents, pd.DataFrame({'name': [f"Опись №{number_inventory}.docx"]})],
                                              ignore_index=True)
                        break
                    else:  # Если не хватает
                        inventory.append(Path(incoming_data['start_path'], f"Опись №{number_inventory}.docx"))
                        documents = pd.concat([documents, pd.DataFrame({'name': [f"Опись №{number_inventory}.docx"]})],
                                              ignore_index=True)
                        number_inventory += 1
            # replace_index = len(order)
            # for index, i in enumerate(order):
            #     if re.findall(r'сопроводит', i.name, re.I):
            #         replace_index = index
            #         break
            # order[replace_index:replace_index] = inventory
        all_doc = len(documents)
        now_doc = 0
        current_progress = 0
        percent = 60 / all_doc
        files_order = sorting_files(start_path, finish_path, False, inventory)
        if files_order['error']:
            if 'trace' in files_order:
                return {'status': 'error', 'text': files_order['text'], 'trace': files_order['trace']}
            else:
                return {'status': 'warning', 'text': files_order['text'], 'trace': ''}
        order = files_order['data']['docs']
        answer = create_text_for_docs(logging, order, documents, line_doing, all_doc, now_doc, progress_value,
                                      line_progress, percent, current_progress, incoming_data)

        if answer['status'] != 'success':
            return answer
        line_progress.emit(f'Выполнено {int(60)} %')
        progress_value.emit(int(60))
        documents = answer['data']['documents']
        documents = documents.reindex(answer['text'])
        incoming_data['secret_number_2'] = answer['data']['num_2']
        documents.reset_index(drop=True, inplace=True)
        # Форма 3
        if incoming_data['flag_inventory'] == 40 and documents["name"].isin(['Форма 3.docx']).any():
            line_doing.emit(f'Генерируем колонтитулы для формы 3')
            logging.info("Генерируем колонтитулы для форму 3")
            index_doc = documents.loc[documents['name'] == 'Форма 3.docx'].index[0]
            documents.loc[index_doc, 'date'] = incoming_data['date']
            documents.loc[index_doc, 'first_header_text'] = f"{incoming_data['classified']}\n" \
                                                            f"{incoming_data['list_item']}\n" \
                                                            f"Экз. №{incoming_data['num_scroll']}"
            documents.loc[index_doc, 'footer_text'] = f"{incoming_data['secret_number_1']}" \
                                                      f"{incoming_data['secret_number_2']}c"
            text_finish = f"Уч. № {documents.loc[index_doc, 'footer_text']}\nОтп. 1 экз. в адрес\n" \
                          f"{incoming_data['hdd_number']}\nИсп. {incoming_data['account_executor']}\n" \
                          f"Печ. {incoming_data['print_executor']}\n{incoming_data['date']}\nб/ч"
            documents.loc[index_doc, 'executor'] = incoming_data['account_executor']
            documents.loc[index_doc, 'text_finish'] = text_finish
            pages = pages_count(documents.loc[index_doc, 'start_path'])
            page = pages['pages']
            if page == 0:
                errors.append(f"Для файла «Форма 3.docx» подсчёт кол-ва страниц завершился с ошибкой: {pages['text']}")
            documents.loc[index_doc, 'pages'] = page
            incoming_data['secret_number_2'] = str(int(incoming_data['secret_number_2']) + 1)  # Увеличиваем значение
        line_progress.emit(f'Выполнено {int(62)} %')
        progress_value.emit(int(62))
        if incoming_data['inventory_insert']:
            line_doing.emit(f'Генерируем колонтитулы для описи(ей)')
            # Новое для преобразования текста для описей.
            a_prescription = documents.loc[documents['a_prescription'].isin([True])]
            if len(a_prescription):
                for item in a_prescription.itertuples():
                    index_doc = documents.loc[documents['name'].str.contains('протокол', case=False)
                                              & documents['number'].str.contains(item.number, case=False)].index[0]
                    a_text = documents.loc[index_doc, 'account_list_text'].split('!')
                    a_text[2] = a_text[2] + '/Приложение несекретно'
                    a_text[3] = a_text[3] + '/' + str(item.pages)
                    documents.loc[index_doc, 'account_list_text'] = '!'.join(a_text)

            logging.info("Добавляем опись")
            answer = inventory_insert(documents, incoming_data, finish_path)
            if answer['status'] != 'success':
                return answer
            documents = answer['data']
        line_progress.emit(f'Выполнено {int(66)} %')
        progress_value.emit(int(66))
        if len(documents[(documents['action'] == 'acc_doc')].index):
            logging.info("Добавляем сопроводительный или запрос")
            dict_file = incoming_data['file_num'] if incoming_data['checkBox_file_num'] else None
            for doc in documents[(documents['action'] == 'acc_doc')].index:
                start_path_acc = documents.loc[doc, 'start_path']
                name_acc = documents.loc[doc, 'name']
                line_doing.emit(f'Генерируем колонтитулы для {name_acc}')
                index_doc = documents.loc[documents['start_path'] == start_path_acc].index[0]
                if dict_file and doc['name'].rpartition('.')[0] in dict_file:  # Если есть файл номеров
                    if re.findall(r'запрос', name_acc, re.I):
                        documents.loc[index_doc, 'footer_text'] = dict_file[name_acc.rpartition('.')[0]][0]
                    else:
                        index_acc = documents.loc[documents['name'].str.contains('сопроводит', case=False)].index[0]
                        documents.loc[index_doc, 'footer_text'] = documents.loc[index_acc, 'footer_text']
                    documents.loc[index_doc, 'date'] = dict_file[name_acc.rpartition('.')[0]][1]  # Дата
                else:
                    if re.findall(r'запрос', name_acc, re.I):
                        documents.loc[index_doc, 'footer_text'] = incoming_data['secret_number_1'] \
                                                                  + incoming_data['secret_number_2'] + 'c'
                    else:
                        index_acc = documents.loc[documents['name'].str.contains('сопроводит', case=False)].index[0]
                        documents.loc[index_doc, 'footer_text'] = documents.loc[index_acc, 'footer_text']
                    documents.loc[index_doc, 'date'] = incoming_data['date']
                documents.loc[index_doc, 'text_finish'] = "Уч. № " + documents.loc[index_doc, 'footer_text'] + \
                                                          "\nОтп. 2 экз.\n№ 1 - в адрес\n№ 2 - в дело \n" \
                                                          + incoming_data['hdd_number'] + "\nИсп. " \
                                                          + incoming_data['executor_acc_sheet'] + "\nПеч. " \
                                                          + incoming_data['print_executor'] + \
                                                          "\n" + incoming_data['date'] + "\nБ/ч"
                pages = pages_count(start_path_acc)
                page = pages['pages']
                if page == 0:
                    errors.append(f"Для файла {name_acc} подсчёт кол-ва страниц завершился с ошибкой: {pages['text']}")
                documents.loc[index_doc, 'pages'] = page
                documents.loc[index_doc, 'account_list'] = False
        line_progress.emit(f'Выполнено {int(70)} %')
        progress_value.emit(int(70))
        if errors:
            return {'status': 'warning', 'data': '', 'text': errors, 'trace': ''}
        return {'status': 'success', 'data': incoming_data, 'documents': documents, 'text': '', 'trace': ''}
    except BaseException as ex:
        return {'status': 'error',
                'text': f'Ошибка при формировании текстовок документов в папке {start_path.name} - {ex}',
                'trace': traceback.format_exc()}


def format_doc(incoming_data: dict, current_progress, now_doc, all_doc, line_doing, line_progress, progress_value,
               event, window_check, info_value) -> dict:
    """Получения данных для заполнения документов и последующее их создание."""
    logging = incoming_data['logging']
    try:
        pt_num = 14 if incoming_data['service'] else 12
        documents = pd.DataFrame()
        if incoming_data['package']:
            folders = [item for item in Path(incoming_data['start_path']).glob('*') if os.path.isdir(item)]
            start_folders = [Path(incoming_data['start_path'], item) for item in folders]
            finish_folders = [Path(incoming_data['finish_path'], item.name) for item in folders]
        else:
            start_folders = [Path(incoming_data['start_path'])]
            finish_folders = [Path(incoming_data['finish_path'])]
        for start_folder, finish_folder in zip(start_folders, finish_folders):
            logging.info(f'Бежим по папке {start_folder.name}')
            answer = add_documents(incoming_data, start_folder, finish_folder, line_doing, line_progress,
                                   progress_value, event, window_check, info_value)
            if answer['status'] != 'success':
                return answer
            incoming_data = answer['data']
            documents = pd.concat([documents, answer['documents']])
            now_doc = 0
            all_doc = len(documents)
            current_progress = 70
            percent = 10 / all_doc
            for document in documents.itertuples():
                line_doing.emit(f'Создаём {document.name} ({now_doc} из {all_doc})')
                logging.info(f'Создаём файл {document.name}')
                # для описи
                account_dict = pd.DataFrame()
                if re.findall(r'опись', document.name.lower()):
                    account_dict = documents.loc[documents['account_list'] == document.name]
                    if '41101' in incoming_data['mode_name']:
                        account_dict.sort_values(by=['number'], ascending=[True], inplace=True, na_position='first')
                if re.findall(r'сопроводит', document.name.lower()):
                    account_dict = documents.loc[documents['name'].str.contains('опись', case=False)]
                answer = create_file(documents, document, pt_num, incoming_data, account_dict)
                if answer['status'] != 'success':
                    return answer
                documents = answer['documents']
                current_progress += percent
                line_progress.emit(f'Выполнено {int(current_progress)} %')
                progress_value.emit(int(current_progress))
                now_doc += 1
            if incoming_data['form27_insert']:
                line_doing.emit(f'Создаём 27 форму')
                logging.info(f'Создаём 27 форму')
                form_27 = documents.loc[documents['form_27'] == 1]
                answer = create_form_27(form_27, finish_folder, incoming_data['form27_firm'])
                if answer['status'] != 'success':
                    return answer
            line_progress.emit(f'Выполнено {int(90)} %')
            progress_value.emit(int(90))
            if incoming_data['main_sp']:
                line_doing.emit(f'Сортируем СП')
                logging.info(f'Сортируем СП')
                name_gk = incoming_data['name_gk'] if incoming_data['checkBox_gk'] else ''
                check_sp = [incoming_data['conclusion_sp'], incoming_data['protocol_sp'],
                            incoming_data['prescription_sp']]
                sp_sorting(name_gk, finish_folder, incoming_data['sp_path_dir'], incoming_data['sp_path_file'],
                           check_sp, progress_value, line_progress)
                if answer['status'] != 'success':
                    return answer
            line_progress.emit(f'Выполнено {int(100)} %')
            progress_value.emit(int(100))
            docs_txt = [file for file in os.listdir(start_folder) if file.endswith('.txt')]  # Список txt
            for txt_file in docs_txt:
                line_doing.emit(f'Копируем txt файлы')
                logging.info(f'Копируем txt файлы')
                shutil.copy(txt_file, finish_folder)
        return {'status': 'success', 'text': '', 'trace': ''}
    except BaseException as error:
        return {'status': 'error', 'text': error, 'trace': traceback.format_exc()}
