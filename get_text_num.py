import traceback
import re
from pathlib import Path
import pandas as pd
from small_functions import pages_count


def create_text_for_docs(log, docs: list, documents: pd.DataFrame, line_doing, all_doc, now_doc,
                         progress_value, line_progress, percent, current_progress,
                         incoming: dict, window_check, event) -> dict:
    """Вставка в основную таблицу номеров, текстовок, дат, исполнителей. Возвращает части секретного номера,
     чтобы продолжить, если включен пакетный режим, а так же список для реиндексации фрейма данных"""
    add_list_item = incoming['add_list_item'] if incoming['checkBox_add_list_item'] else ''
    dict_file = incoming['file_num'] if incoming['checkBox_file_num'] else {}
    # Имена чтобы не увеличивался секретный номер
    doc_name_for_continue = ['result', 'инфокарта', 'приложение а']
    doc_name_for_dict_40 = ['акт', 'заключение', 'протокол', 'предписание', 'утверждение', 'опись']
    reindex_list = []
    doc_for_log = ''
    try:
        errors = []
        for doc in docs:  # Для файлов в папке
            event.wait()
            if window_check.stop_threading:
                return {'status': 'cancel', 'trace': '', 'text': ''}
            doc_name = doc.name
            doc_for_log = doc.name
            parent_path = doc.parent
            line_doing.emit(f'Генерируем колонтитулы для {doc_name} ({now_doc} из {all_doc})')
            # log.info(f"Добавляем данные для {doc_name}")
            if re.findall(r'опись', doc_name, re.I):
                index_doc = documents.loc[documents['name'] == doc.name].index[0]
            else:
                index_doc = documents.loc[documents['start_path'] == doc].index[0]
            reindex_list.append(index_doc)
            # Определяем есть ли номер. Сначала отсекаем, потом ищем структуру.
            number_doc = doc_name.rpartition('.')[0].rpartition(' ')[2]
            if not re.findall(r'\d\.\d', number_doc):
                number_doc = '0.0'
            text_first_header = f"{incoming['classified']}\n{incoming['list_item']}\nЭкз. №{incoming['num_scroll']}"
            footer_text = False
            date = incoming['date']
            executor = False
            if dict_file and doc_name.rpartition('.')[0] in dict_file:  # Если есть файл номеров
                footer_text = dict_file[doc_name.rpartition('.')[0]][0]  # Текст для нижнего колонтитула
                date = dict_file[doc_name.rpartition('.')[0]][1]  # Дата
            else:
                if all([True if _ not in doc_name.lower() else False for _ in doc_name_for_continue]):
                    footer_text = incoming['secret_number_1'] + incoming['secret_number_2'] + 'c'  # Нижний колонтитул
            if doc_name.lower() == 'форма 3.docx':
                executor = incoming['print_executor']
            elif 'result' in doc_name.lower():
                documents.loc[index_doc, 'action'] = 'copy'
                current_progress += percent
                line_progress.emit(f'Выполнено {int(current_progress)} %')
                progress_value.emit(int(current_progress))
                now_doc += 1
                continue
            elif re.findall('сопроводит', doc_name.lower()):
                documents.loc[index_doc, 'action'] = 'acc_doc'
                executor = incoming['executor_acc_sheet']
            elif re.findall('запрос', doc_name.lower()):
                documents.loc[index_doc, 'action'] = 'acc_doc'
                executor = incoming['executor_acc_sheet']
            elif re.findall('инфокарта', doc_name.lower()):
                documents.loc[index_doc, 'action'] = 'copy'
                # documents.loc[index_doc, 'number'] = number_doc
            elif re.findall(r'приложение а', doc_name.lower()):
                # documents.loc[index_doc, 'number'] = number_doc
                documents.loc[index_doc, 'a_prescription'] = True
                protocol_name = documents[(documents['name'].str.contains(r'протокол', case=False))
                                          & (documents['number'] == number_doc)
                                          & (documents['parent_path'] == parent_path)]
                protocol_name = protocol_name.reset_index(drop=True)
                if len(protocol_name) > 0:
                    documents.loc[
                        index_doc, 'text'] = f"к протоколу уч. № {str(protocol_name.loc[0, 'footer_text'])} от date"
                    documents.loc[index_doc, 'change_date'] = True
                else:
                    errors.append(f"Для файла {doc_name} не нашли протокол, документ не заполнен")
            elif re.findall(r'приложение', doc_name.lower()):
                if re.findall(r'заключени', doc_name.lower()):
                    conclusion_name = documents[(documents['name'].str.contains(r'заключение', case=False))]
                    if len(conclusion_name) == 0:
                        documents.loc[index_doc, 'text'] = f"от {incoming['conclusion_number_date']}" \
                                                           f" № {str(incoming['conclusion_number'])}"
                    elif len(conclusion_name) == 1:
                        documents.loc[index_doc, 'text'] = f"от {incoming['conclusion_number_date']}" \
                                                           f" № {str(incoming['conclusion_number'])}"
                    else:
                        # Такого случая не предусмотрено, если произошло - косяк.
                        log.warning('НЕСТАНДАРТНАЯ СИТУАЦИЯ, АЛГОРИТМ НЕ ПРОДУМАН И НЕ ОТЛАЖЕН')
                        errors.append(f'В {doc_name} не добавлен секретный номер, ситуация не согласована')
                        documents.loc[index_doc, 'text'] = f"от {date} № "
                    executor = incoming['protocol']
                else:
                    act_number = documents[documents['name'].str.contains(r'акт', case=False)]
                    index_act = act_number.index.tolist()[0]
                    documents.loc[index_doc, 'text'] = f"от date № {act_number.loc[index_act, 'footer_text']}"
                    executor = incoming['act_executor']
                documents.loc[index_doc, 'change_date'] = True
                if add_list_item:
                    text_first_header = incoming['classified'] + '\n' + add_list_item + '\nЭкз. №' + incoming[
                        'num_scroll']
            elif re.findall(r'заключение', doc_name.lower()):
                # documents.loc[index_doc, 'number'] = number_doc
                executor = incoming['conclusion']
                documents.loc[index_doc, 'change_date'] = True
            elif re.findall(r'протокол', doc_name.lower()):
                # Параграф для колонтитула первой страницы
                if add_list_item:
                    text_first_header = incoming['classified'] + '\n' + add_list_item + '\nЭкз. №' + incoming[
                        'num_scroll']
                # documents.loc[index_doc, 'number'] = number_doc
                documents.loc[index_doc, 'text_conclusion'] = ''
                conclusion_name = documents[(documents['name'].str.contains(r'заключение', case=False)
                                             & (documents['parent_path'] == parent_path))]
                if len(conclusion_name) > 1:
                    conclusion_name = documents[(documents['name'].str.contains(r'заключение', case=False)
                                                 & (documents['number'] == number_doc)
                                                 & (documents['parent_path'] == parent_path))]
                conclusion_name = conclusion_name.reset_index(drop=True)
                if len(conclusion_name) == 0:
                    if incoming['checkBox_conclusion_number']:
                        documents.loc[
                            index_doc, 'text_conclusion'] = f"уч. № {str(incoming['conclusion_number'])} от " \
                                                            f"{incoming['conclusion_number_date']}"
                else:
                    documents.loc[
                        index_doc, 'text_conclusion'] = f"уч. № {str(conclusion_name.loc[0, 'footer_text'])} от" \
                                                        f" {conclusion_name.loc[0, 'date']}"
                executor = incoming['protocol']
                documents.loc[index_doc, 'change_date'] = True
            elif re.findall(r'предписание', doc_name.lower()):
                # Параграф для колонтитула первой страницы
                if add_list_item:
                    text_first_header = incoming['classified'] + '\n' + add_list_item + '\nЭкз. №' + incoming[
                        'num_scroll']
                # documents.loc[index_doc, 'number'] = number_doc
                documents.loc[index_doc, 'text_conclusion'] = ''
                documents.loc[index_doc, 'text_protocol'] = ''
                conclusion_name = documents[(documents['name'].str.contains(r'заключение', case=False)
                                             & (documents['parent_path'] == parent_path))]
                if len(conclusion_name) > 1:
                    conclusion_name = documents[(documents['name'].str.contains(r'заключение', case=False)
                                                 & (documents['number'] == number_doc)
                                                 & (documents['parent_path'] == parent_path))]
                protocol_name = documents[(documents['name'].str.contains(r'протокол', case=False)
                                           & (documents['number'] == number_doc)
                                           & (documents['parent_path'] == parent_path))]
                conclusion_name = conclusion_name.reset_index(drop=True)
                protocol_name = protocol_name.reset_index(drop=True)
                if len(protocol_name):
                    documents.loc[
                        index_doc, 'text_protocol'] = f"уч. № {str(protocol_name.loc[0, 'footer_text'])} от" \
                                                      f" {protocol_name.loc[0, 'date']}"
                else:
                    if not incoming['dont_check']:
                        errors.append(f"Для файла {doc_name} не нашли протокол, документ не заполнен")
                if len(conclusion_name) == 0:
                    if incoming['checkBox_conclusion_number']:
                        documents.loc[
                            index_doc, 'text_conclusion'] = f"уч. № {str(incoming['conclusion_number'])} от" \
                                                            f" {incoming['conclusion_number_date']}"
                else:
                    documents.loc[
                        index_doc, 'text_conclusion'] = f"уч. № {str(conclusion_name.loc[0, 'footer_text'])} от" \
                                                        f" {conclusion_name.loc[0, 'date']}"
                executor = incoming['prescription']
                documents.loc[index_doc, 'change_date'] = True
            elif re.findall(r'акт', doc_name.lower()):
                executor = incoming['act_executor']
                documents.loc[index_doc, 'change_date'] = True
            elif re.findall(r'утверждение', doc_name.lower()):
                executor = incoming['statement_executor']
                documents.loc[index_doc, 'change_date'] = True
            elif re.findall(r'опись', doc_name.lower()):
                executor = incoming['account_executor']
            else:
                # log.info('Документ не в списке необходимых, продолжаем')
                current_progress += percent
                line_progress.emit(f'Выполнено {int(current_progress)} %')
                progress_value.emit(int(current_progress))
                now_doc += 1
                continue
            text_finish = f"Уч. № {footer_text}\nОтп. 1 экз. в адрес\n{incoming['hdd_number']}\nИсп. {executor}\n" \
                          f"Тел. {incoming['telephone']}\nПеч. {incoming['print_executor']}\n{date}\nб/ч"
            documents.loc[index_doc, 'number'] = number_doc
            documents.loc[index_doc, 'date'] = date
            documents.loc[index_doc, 'first_header_text'] = text_first_header
            documents.loc[index_doc, 'footer_text'] = footer_text
            documents.loc[index_doc, 'text_finish'] = text_finish
            documents.loc[index_doc, 'executor'] = executor
            documents.loc[index_doc, 'form_27'] = 1
            if any([True if _ in doc_name.lower() else False for _ in doc_name_for_continue]):
                documents.loc[index_doc, 'form_27'] = 0
            if re.findall(r'опись', doc_name.lower()):
                documents.loc[index_doc, 'num_scroll'] = 1
                if not dict_file:
                    incoming['secret_number_2'] = str(int(incoming['secret_number_2']) + 1)  # Увеличиваем учетный номер
                current_progress += percent
                line_progress.emit(f'Выполнено {int(current_progress)} %')
                progress_value.emit(int(current_progress))
                now_doc += 1
                continue
            if re.findall(r'приложение а', doc_name.lower()):
                pages = pages_count(doc)
                page = pages['pages']
                if page == 0:
                    errors.append(f"Для файла {doc_name} подсчёт кол-ва страниц завершился с ошибкой: {pages['text']}")
                documents.loc[index_doc, 'pages'] = page
            documents.loc[index_doc, 'account_list'] = 0
            for dict_40_name in doc_name_for_dict_40:
                if dict_40_name in doc_name.lower():
                    documents.loc[index_doc, 'account_list'] = 1
                    # documents.loc[index_doc, 'account_list_text'] = f"{doc_name[:-5]}" \
                    #                                                 f"!{footer_text}" \
                    #                                                 f"!№{documents.loc[index_doc, 'num_scroll']}," \
                    #                                                 f" {incoming['classified']}" \
                    #                                                 f"!{page}"
                    documents.loc[index_doc, 'account_list_text'] = f"{doc_name[:-5]}" \
                                                                    f"!{footer_text}" \
                                                                    f"!№{documents.loc[index_doc, 'num_scroll']}," \
                                                                    f" {incoming['classified']}" \
                                                                    f"!page"
            # Добавляем 2 сопроводительный, если это сопроводительный :)
            if re.findall('сопроводит', doc_name.lower()):
                documents.loc[
                    index_doc, 'text_finish'] = f"Исполнил {executor}\nТелефон {incoming['telephone_acc_sheet']}" \
                                                f"\nДобавочный номер {incoming['add_telephone']}"
                documents = pd.concat([documents, documents[documents['start_path'] == doc]], ignore_index=True)
                index_doc_acc = len(documents) - 1
                documents.loc[
                    index_doc_acc, 'name'] = f"{documents.loc[index_doc_acc, 'name'].rpartition('.')[0]} (2 экз.).docx"
                documents.loc[index_doc_acc, 'finish_path'] = Path(documents.loc[index_doc_acc, 'finish_path'].parent,
                                                                   documents.loc[index_doc_acc, 'name'])
                documents.loc[index_doc_acc, 'first_header_text'] = f"{incoming['classified']}" \
                                                                    f"\n{incoming['list_item']}\nЭкз.№2"
                reindex_list.append(index_doc_acc)
            if 'second_copy' in incoming and incoming['second_copy']:
                max_copy = incoming['number_instance'][len(incoming['number_instance']) - 1]
                for number_folder in incoming['number_instance']:
                    for index, (value1, value2) in enumerate(
                            zip(incoming['second_copy'], ['заключение', 'протокол', 'предписание'])):
                        if value1 and re.findall(value2, doc_name.lower()):
                            copy_doc = documents.loc[index_doc].copy()
                            documents.loc[
                                index_doc, 'text_finish'] = f"Уч. № {footer_text}\nОтп. {str(max_copy)}" \
                                                            f" экз. в адрес\n{incoming['hdd_number']}" \
                                                            f"\nИсп. {executor}\nТел. {incoming['telephone']}" \
                                                            f"\nПеч. {incoming['print_executor']}\n{date}\nб/ч"
                            documents.loc[len(documents)] = copy_doc
                            index_doc_new = len(documents) - 1
                            documents.loc[index_doc_new, 'first_header_text'] = incoming['classified'] + '\n'\
                                + incoming['list_item'] + '\nЭкз. №' + str(number_folder)
                            documents.loc[index_doc_new, 'finish_path'] = Path(
                                documents.loc[index_doc, 'finish_path'].parent,
                                f"{str(number_folder)} экземпляр",
                                documents.loc[index_doc_new, 'finish_path'].name)
                            documents.loc[index_doc_new, 'parent_path'] = Path(
                                documents.loc[index_doc, 'finish_path'].parent)
                            documents.loc[index_doc_new, 'num_scroll'] = str(number_folder)
                            documents.loc[
                                index_doc_new, 'text_finish'] = f"Уч. № {footer_text}\nОтп. {str(max_copy)}" \
                                                                f" экз. в адрес\n{incoming['hdd_number']}" \
                                                                f"\nИсп. {executor}\nТел. {incoming['telephone']}" \
                                                                f"\nПеч. {incoming['print_executor']}\n{date}\nб/ч"
                            # documents.loc[index_doc_new, 'account_list_text'] = f"{doc_name[:-5]}" \
                            #                                                     f"!{footer_text}" \
                            #                                                     f"!№{number_folder}," \
                            #                                                     f" {incoming['classified']}" \
                            #                                                     f"!{page}"
                            documents.loc[index_doc_new, 'account_list_text'] = f"{doc_name[:-5]}" \
                                                                                f"!{footer_text}" \
                                                                                f"!№{number_folder}," \
                                                                                f" {incoming['classified']}" \
                                                                                f"!main_page"
                            reindex_list.append(index_doc_new)
                            break
            if not dict_file and all([True if _ not in doc_name.lower() else False for _ in doc_name_for_continue]):
                incoming['secret_number_2'] = str(int(incoming['secret_number_2']) + 1)  # Увеличиваем учетный номер
            current_progress += percent
            line_progress.emit(f'Выполнено {int(current_progress)} %')
            progress_value.emit(int(current_progress))
            now_doc += 1
        if errors:
            return {'status': 'warning', 'text': errors, 'data': '', 'trace': ''}
        else:
            return {'status': 'success', 'text': reindex_list, 'trace': '',
                    'data': {'num_2': incoming['secret_number_2'], 'documents': documents}}
    except BaseException as exception:
        log.error(f'Имя документа - {doc_for_log}')
        return {'status': 'error', 'text': str(exception), 'trace': traceback.format_exc()}
