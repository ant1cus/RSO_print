import os
import re

import openpyxl
import pandas as pd
import psutil
from pathlib import Path


def check(n, e):
    for symbol in e:
        if n == symbol:
            return False
    return True


def check_doc_format(incoming: dict) -> dict:
    for proc in psutil.process_iter():
        if proc.name() == 'WINWORD.EXE':
            return {'error': True, 'data': 'Закройте все файлы Word!'}
    # Путь к исходным документам и проверки
    if not incoming['start_path']:
        return {'error': True, 'data': 'Путь к исходным документам пуст'}
    if os.path.exists(incoming['start_path']) is False:
        return {'error': True, 'data': 'Папка с исходными документами отсутствует или переименована'}
    if os.path.isfile(incoming['start_path']):
        return {'error': True, 'data': 'В исходных документах указан файл, а не папка'}
    if len(os.listdir(incoming['start_path'])) == 0:
        return {'error': True, 'data': 'Папка с исходными документами пуста'}
    if incoming['package']:
        docs = []
        folders = [i for i in os.listdir(incoming['start_path']) if os.path.isdir(Path(incoming['start_path'], i))]
        if len(folders) == 0:
            return {'error': True, 'data': 'В указанной директории нет ни одной папки работы в пакетном режиме'}
        for folder in os.listdir(incoming['start_path']):
            if os.path.isdir(incoming['start_path'] + '\\' + folder):
                # Ошибка если есть файлы старого формата
                docs += [i for i in os.listdir(incoming['start_path'] + '\\' + folder) if i[-3:] == 'doc']
    else:
        error = [i for i in os.listdir(incoming['start_path']) if os.path.isdir(incoming['start_path'] + '\\' + i) and
                 ('материалы' not in i.lower())]
        if error:
            return {'error': True, 'data': 'В директории для преобразования присутствуют папки'}
        # Ошибка если есть файлы старого формата
        docs = [i for i in os.listdir(incoming['start_path']) if i[-3:] == 'doc']
    if len(docs) != 0:
        text = 'Файлы старого формата:\n' + '\n'.join(docs)
        return {'error': True, 'data': text}
    if incoming['action_mo']:
        if incoming['package']:
            for folder_mo in os.listdir(incoming['start_path']):
                file_mo = [mo for mo in os.listdir(incoming['start_path'] + '\\' + folder_mo)
                           if mo[-3:] == 'txt' and 'F19' in mo]
                if not file_mo:
                    return {'error': True, 'data': f'Нет текстового файла с серийниками для создания отчёта для'
                                                   f' МО в папке {folder_mo}'}
        else:
            file_mo = [mo for mo in os.listdir(incoming['start_path']) if mo[-3:] == 'txt' and 'F19' in mo]
            if not file_mo:
                return {'error': True, 'data': 'Нет текстового файла с серийниками для создания отчёта для МО'}
    # Путь к конечным документам и проверки
    if not incoming['finish_path']:
        return {'error': True, 'data': 'Путь к конечной папке пуст'}
    if os.path.exists(incoming['finish_path']) is False:
        return {'error': True, 'data': 'Конечная папка отсутствует или переименована'}
    if os.path.isfile(incoming['finish_path']):
        return {'error': True, 'data': 'Указанный путь к конечной папке не является директорией'}
    if len(os.listdir(incoming['finish_path'])) != 0:
        return {'error': True, 'data': 'Конечная папка не пуста, очистите директорию'}
    if incoming['checkBox_file_num']:
        if not incoming['file_num']:
            return {'error': True, 'data': 'Не указан файл номеров'}
        if os.path.isdir(incoming['file_num']):
            return {'error': True, 'data': 'Указанный путь к файлу номеров является директорией'}
        else:
            if os.path.exists(incoming['file_num']):
                if incoming['file_num'].endswith('.xlsx') is False:
                    return {'error': True, 'data': 'Файл номеров не формата .xlsx'}
            else:
                return {'error': True, 'data': 'Файл номеров удалён или переименован'}
        file_in_directory = []
        if incoming['package']:
            for directory in os.listdir(incoming['start_path']):
                file_in_directory = [file for file in os.listdir(Path(incoming['start_path'], directory))
                                     if file.endswith('.docx')]
        else:
            file_in_directory = [file for file in os.listdir(incoming['start_path']) if file.endswith('.docx')]
        dict_file = {}
        wb = openpyxl.load_workbook(incoming['file_num'])  # Откроем книгу.
        ws = wb.active  # Делаем активным первый лист.
        error = []
        for i in range(1, ws.max_row + 1):  # Пока есть значения
            if ws.cell(i, 1).value:
                try:
                    check_date = ws.cell(i, 3).value.strftime("%d.%m.%Y")
                except AttributeError:
                    check_date = ''
                if ws.cell(i, 2).value is None or len(ws.cell(i, 2).value) == 0:
                    error.append(f'В строке {i} файла excel нет секретного номера')
                else:
                    dict_file[ws.cell(i, 1).value] = [ws.cell(i, 2).value + 'c', check_date]  # Делаем список
        if error:
            return {'error': True, 'data': '\n'.join(error)}
        error_for_file_num = []
        # for file in file_in_directory:
        #     accepted_file = [False if name_file in file.lower() else True for name_file in ['акт', 'заключение', 'протокол', 'предписание', 'сопроводит', 'опись']]
        #     if all(accepted_file):
        #         continue
        #     num_date = dict_file.pop(file.rpartition('.')[0], 'File not found')
        #     if num_date == 'File not found':
        #         error_for_file_num.append(f'Документ {file} не найден в файле номеров')
        #     else:
        #         if num_date[0] is False:
        #             error_for_file_num.append(f'Для записи {file} в файле номеров не указан секретный номер')
        #         elif num_date[1] is False:
        #             error_for_file_num.append(f'Для записи {file} в файле номеров не указана дата')
        # if dict_file:
        for file in dict_file:
            if len(dict_file[file][0]) == 0:
                error_for_file_num.append(f'Для записи {file} в файле номеров не указан секретный номер')
            elif len(dict_file[file][1]) == 0:
                error_for_file_num.append(f'Для записи {file} в файле номеров не указана дата')
        if error_for_file_num:
            return {'error': True, 'data': '\n'.join(error_for_file_num)}
        incoming['file_num'] = dict_file
    if incoming['checkBox_signature']:
        if not incoming['path_signature']:
            return {'error': True, 'data': 'Путь к папке с подписями пуст'}
        if os.path.exists(incoming['path_signature']) is False:
            return {'error': True, 'data': 'Папка с подписями отсутствует или переименована'}
        if os.path.isfile(incoming['path_signature']):
            return {'error': True, 'data': 'Указанный путь к папке с подписями не является директорией'}
        # if os.listdir(incoming['path_signature']) is False:
        #     return {'error': True, 'data': 'В указанная папка с подписями пуста'}
        incoming['signature_files'] = {}
        for enum, file in enumerate(os.listdir(incoming['path_signature'])):
            if os.path.isfile(Path(incoming['path_signature'], file)) and re.findall(r'[A-z]*', file):
                incoming['signature_files'][file.partition('.')[0]] = str(Path(incoming['path_signature'], file))
        if not len(incoming['signature_files']):
            return {'error': True, 'data': 'В указанной папке с подписями нет подходящих файлов'}
    if 'main_sp' in incoming and incoming['main_sp']:
        if not incoming['sp_path_dir']:
            return {'error': True, 'data': 'Путь к папке с материалами СП пуст'}
        if os.path.isfile(incoming['sp_path_dir']):
            return {'error': True, 'data': 'Указанный путь к материалам СП не является директорией'}
        if not incoming['sp_path_file']:
            return {'error': True, 'data': 'Путь к файлу с номерами СП пуст'}
        if os.path.isdir(incoming['sp_path_file']):
            return {'error': True, 'data': 'Указанный путь к файлу с номерами СП не является файлом'}
        if os.path.exists(incoming['sp_path_file']) is False:
            return {'error': True, 'data': 'Файл номеров СП удалён или переименован'}
        if incoming['sp_path_file'].endswith('.xlsx') is False:
            return {'error': True, 'data': 'Файл номеров СП не формата .xlsx'}
        if incoming['checkBox_gk']:
            if not incoming['name_gk']:
                return {'error': True, 'data': 'Введите имя ГК'}
        if len(os.listdir(incoming['sp_path_dir'])) == 0 and incoming['name_gk'] is False:
            return {'error': True, 'data': 'Папка с материалами СП пуста (введите имя ГК или добавьте материалы)'}
        incoming['check_sp'] = [True if i else False for i in [incoming['conclusion_sp'], incoming['protocol_sp'],
                                                               incoming['prescription_sp'], incoming['infocard_sp']
                                                               ]
                                ]
        if all(i is False for i in incoming['check_sp']):
            return {'error': True, 'data': 'Не выбран ни один документ для проверки СП'}
    # Ведомство
    if incoming['radioButton_FSB']:
        incoming['service'] = True
    elif incoming['radioButton_FSTEK']:
        incoming['service'] = False
    else:
        return {'error': True, 'data': 'Не выбрано ведомство для вставки колонтитулов'}
    # Гриф секретности
    class_ = {'ДСП': 'Для служебного пользования', 'С': 'Секретно', 'СС': 'Совершенно секретно',
              'ОВ': 'Особой важности'}
    if not incoming['classified']:
        return {'error': True, 'data': 'Не выбрана категория секретности'}
    incoming['classified'] = class_[incoming['classified']]
    # Номер экземпляра
    if not incoming['num_scroll']:
        return {'error': True, 'data': 'Не указан номер экземпляра'}
    # Пункт перечня
    if not incoming['list_item']:
        return {'error': True, 'data': 'Не указан пункт перечня'}
        # Дополнительный пункт перечня
    if incoming['checkBox_add_list_item']:
        if not incoming['add_list_item']:
            return {'error': True, 'data': 'Не указан дополнительный пункт перечня'}
    # Номер
    incoming['secret_number_1'] = ''
    incoming['secret_number_2'] = ''
    if incoming['checkBox_file_num'] is False:
        if incoming['number']:
            if incoming['number'][-1] in ['С', 'с']:
                incoming['number'] = incoming['number'].replace(incoming['number'][-1], 'c')
            if not incoming['number']:
                return {'error': True, 'data': 'Не указан номер'}
            for i in incoming['number']:
                if check(i, ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0', '/', 'c', 'с', '-', 'Н', 'С', 'с')):
                    return {'error': True, 'data': 'Есть лишние символы в номере'}
            if (re.match(r'\w+/\w+/\w+c$', incoming['number']) is None) and (re.match(r'НС-\w+c$',
                                                                                      incoming['number']) is None):
                return {'error': True, 'data': 'Секретный номер указан неверно'}
            if re.match(r'\w+/\w+/\w+c', incoming['number']):
                incoming['secret_number_1'] = incoming['number'].rpartition('/')[0] + '/'
                incoming['secret_number_2'] = incoming['number'].rpartition('/')[2].rpartition('c')[0]
            else:
                incoming['secret_number_1'] = incoming['number'].partition('-')[0] + '-'
                incoming['secret_number_2'] = incoming['number'].partition('-')[2].rpartition('c')[0]
        else:
            return {'error': True, 'data': 'Не указан секретный номер'}
    # Исполнитель, заключение, предписание, протокол, печать
    list_label = ['Исп. заключения', 'Исп. протокола', 'Исп. предписания', 'Исп. печать', 'Исп. сопроводит.']
    for i, element in enumerate([incoming['conclusion'], incoming['protocol'], incoming['prescription'],
                                 incoming['print_executor'], incoming['executor_acc_sheet']]):
        if not element:
            return {'error': True, 'data': 'Не указан(а) ' + list_label[i]}
    if incoming['checkBox_conclusion_number']:
        if incoming['conclusion_number'][-1] in ['С', 'с']:
            incoming['conclusion_number'] = incoming['conclusion_number'].replace(incoming['conclusion_number'][-1], 'c')
        if not incoming['conclusion_number']:
            return {'error': True, 'data': 'Не указан номер заключения'}
        for i in incoming['conclusion_number']:
            if check(i, ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0', '/', 'c', 'с', '-', 'Н', 'С', 'с')):
                return {'error': True, 'data': 'Есть лишние символы в номере заключения'}
        if (re.match(r'\w+/\w+/\w+c$', incoming['conclusion_number']) is None)\
                and (re.match(r'НС-\w+c$', incoming['conclusion_number']) is None):
            return {'error': True, 'data': 'Номер заключения указан неверно'}
    # answer['account'], answer['flag_inventory'], answer['account_post'] = False, False, False
    # answer['account_signature'], answer['account_path'] = False, False
    if incoming['inventory_insert']:
        if incoming['radioButton_40_num'] or incoming['radioButton_all_doc']:
            incoming['flag_inventory'] = 40 if incoming['radioButton_40_num'] else 1
        else:
            return {'error': True, 'data': 'Не указано количество документов в описи'}
        if not incoming['account_position']:
            return {'error': True, 'data': 'Не указана должность для описи'}
        if not incoming['account_executor']:
            return {'error': True, 'data': 'Не указана подпись для описи'}
        if not incoming['account_path']:
            return {'error': True, 'data': 'Не указан путь для файла описи'}
        else:
            if os.path.isfile(incoming['account_path']):
                return {'error': True, 'data': 'Для описи необходимо указать директорию'}
    if not incoming['hdd_number']:
        return {'error': True, 'data': 'Отсутствует номер жесткого диска'}
    if incoming['form27_insert']:
        if not incoming['form27_firm']:
            return {'error': True, 'data': 'Не заполнена организация для 27 формы'}
        if not incoming['form27_path']:
            return {'error': True, 'data': 'Нет пути для 27 формы'}
        if os.path.isfile(incoming['form27_path']):
            return {'error': True, 'data': 'Указанный путь для 27 формы не является директорией'}
    incoming['second_copy'] = []
    if 'main_instance' in incoming and incoming['main_instance']:
        if not incoming['number_instance']:
            return {'error': True, 'data': 'Не указаны номера экземпляров'}
        for i in incoming['number_instance']:
            if check(i, ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0', ' ', '-', ',', '.')):
                return {'error': True, 'data': 'Есть лишние символы в номерах экземпляров'}
        set_num = incoming['number_instance'].replace(' ', '').replace(',', '.')
        if set_num[0] == '.' or set_num[0] == '-':
            return {'error': True, 'data': 'Первый символ номера экземпляра введён не верно'}
        if set_num[-1] == '.' or set_num[-1] == '-':
            return {'error': True, 'data': 'Последний символ номера экземпляра введён не верно'}
        for i in range(len(set_num)):
            if set_num[i] == '.' or set_num[i] == '-':
                if set_num[i + 1] == '.' or set_num[i + 1] == '-':
                    return {'error': True, 'data': 'Два разделителя номеров экземпляра подряд'}
        set_number = []
        for element in set_num.split('.'):
            if '-' in element:
                num1, num2 = int(element.partition('-')[0]), int(element.partition('-')[2])
                if num1 >= num2:
                    return {'error': True, 'data': 'Диапазон номеров экземпляров указан не верно'}
                else:
                    for el in range(num1, num2 + 1):
                        set_number.append(el)
            else:
                set_number.append(element)
        set_number.sort()
        if len(set_number) != len(set(set_number)):
            return {'error': True, 'data': 'Есть повторения в номерах экземпляров'}
        incoming['number_instance'] = set_number
        incoming['second_copy'] = [incoming['conclusion_instance'], incoming['protocol_instance'],
                                   incoming['prescription_instance']]
        if all(i is False for i in incoming['second_copy']):
            return {'error': True, 'data': 'Не выбран ни один документ для создания экземпляров'}
    return {'error': False, 'data': incoming}


def check_doc_print(incoming: dict) -> dict:
    # Ведомство
    if incoming['fsb']:
        incoming['service'] = True
    elif incoming['fstek']:
        incoming['service'] = False
    else:
        return {'error': True, 'data': 'Не выбрано ведомство при печати документов'}
    if not incoming['start_path']:
        return {'error': True, 'data': 'Путь к исходным документам для печати пуст'}
    if os.path.isfile(incoming['start_path']):
        return {'error': True, 'data': 'Путь к исходным документам для печати отсутствует или переименован'}
    if len(os.listdir(incoming['start_path'])) == 0:
        return {'error': True, 'data': 'Папка с исходными документами для печати пуста'}
    docs = []
    if incoming['package']:
        error = [i for i in os.listdir(incoming['start_path'])
                 if os.path.isfile(Path(incoming['start_path'], i))]
        if error:
            return {'error': True, 'data': 'В директории для пакетной печати присутствуют файлы'}
        for folder in os.listdir(incoming['start_path']):
            # Ошибка если есть файлы старого формата
            docs = docs + [i for i in os.listdir(Path(incoming['start_path'], folder)) if i[-3:] == 'doc']
    else:
        error = [i for i in os.listdir(incoming['start_path'])
                 if os.path.isdir(Path(incoming['start_path'], i))]
        if error:
            return {'error': True, 'data': 'В директории для преобразования присутствуют папки'}
        # Ошибка если есть файлы старого формата
        docs = [i for i in os.listdir(incoming['start_path']) if i[-3:] == 'doc']
    if len(docs) != 0:
        text = 'Файлы старого формата:\n' + '\n'.join(docs)
        return {'error': True, 'data': text}
    # Путь к номерам
    if incoming['check_box_add_account_num']:
        if not incoming['add_path_account_num']:
            return {'error': True, 'data': 'Путь к доп. файлу номеров учетных листов пуст'}
        if os.path.isdir(incoming['add_path_account_num']):
            return {'error': True, 'data': 'Указанный путь к доп. файлу номеров учётных листов является директорией'}
        else:
            if os.path.exists(incoming['add_path_account_num']):
                if incoming['add_path_account_num'].endswith('.xlsx') is False:
                    return {'error': True, 'data': 'Доп. файл номеров не формата .xlsx'}
            else:
                return {'error': True, 'data': 'Доп. файл номеров удалён или переименован'}
    if not incoming['path_account_num']:
        return {'error': True, 'data': 'Путь к файлу номеров учетных листов пуст'}
    if os.path.isdir(incoming['path_account_num']):
        return {'error': True, 'data': 'Указанный путь к файлу номеров учётных листов является директорией'}
    if os.path.exists(incoming['path_account_num']):
        if incoming['path_account_num'].endswith('.xlsx'):
            try:
                df_acc_num = pd.read_excel(incoming['path_account_num'], header=None)
                if df_acc_num.empty:
                    return {'error': True, 'data': 'Файл номеров пустой'}
            except BaseException:
                return {'error': True, 'data': 'Что-то не так с файлом номеров'}
        else:
            return {'error': True, 'data': 'Файл номеров не формата .xlsx'}
    else:
        return {'error': True, 'data': 'Файл номеров удалён или переименован'}
    # Форма 27
    if incoming['check_box_from_27']:
        if incoming['package'] is False:
            if not incoming['path_form_27']:
                return {'error': True, 'data': 'Путь к 27 форме пуст'}
            if os.path.isdir(incoming['path_form_27']):
                return {'error': True, 'data': 'Указанный путь к 27 форме является директорией'}
            else:
                if os.path.exists(incoming['path_form_27']) and not incoming['package']:
                    if incoming['path_form_27'].endswith('.xlsx') is False:
                        return {'error': True, 'data': 'Файл "Форма 27" не формата .xlsx'}
                else:
                    if not incoming['package']:
                        return {'error': True, 'data': 'Файл "Форма 27" удалён или переименован'}
        else:
            incoming['path_form_27'] = True
    # Способ печати
    if incoming['duplex'] is False and incoming['last_duplex'] is False and incoming['one_side'] is False:
        return {'error': True, 'data': 'Не указан метод печати'}
    if not incoming['name_printer']:
        return {'error': True, 'data': 'Не выбран принтер'}
    return {'error': False, 'data': incoming}


def check_create_instance_number(incoming: dict) -> dict:
    start_path = incoming['start_path']
    if not start_path:
        return {'error': True, 'data': 'Путь к исходным экземплярам документов пуст'}
    if not os.path.isdir(start_path):
        return {'error': True, 'data': 'Указанный путь к исходным экземплярам документов не является директорией'}
    files = 0
    for file in Path(start_path).rglob('*.*'):
        if file.suffix != '.docx' :
            continue
        for f in ['заключение', 'протокол', 'предписание']:
            if f in file.name.lower():
                files += 1
    finish_path = incoming['finish_path']
    if not finish_path:
        return {'error': True, 'data': 'Путь к конечным экземплярам документов пуст'}
    if not os.path.isdir(finish_path):
        return {'error': True, 'data': 'Указанный путь к конечным экземплярам документов не является директорией'}
    number_instance = incoming['number_instance']
    for i in number_instance:
        if check(i, ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0', ' ', '-', ',', '.')):
            return {'error': True, 'data': 'Есть лишние символы в номерах экземпляров документов'}
    set_num = number_instance.replace(' ', '').replace(',', '.')
    if set_num[0] == '.' or set_num[0] == '-':
        return {'error': True, 'data': 'Первый символ в номерах экземпляра документов введён не верно'}
    if set_num[-1] == '.' or set_num[-1] == '-':
        return {'error': True, 'data': 'Последний символ в номерах экземпляра документов введён не верно'}
    for i in range(len(set_num)):
        if set_num[i] == '.' or set_num[i] == '-':
            if set_num[i + 1] == '.' or set_num[i + 1] == '-':
                return {'error': True, 'data': 'Два разделителя номеров подряд в номерах экземпляра документов'}
    set_number = []
    for element in set_num.split('.'):
        if '-' in element:
            num1, num2 = int(element.partition('-')[0]), int(element.partition('-')[2])
            if num1 >= num2:
                return {'error': True, 'data': 'Диапазон номеров экземпляров документов указан неверно'}
            else:
                for el in range(num1, num2 + 1):
                    set_number.append(el)
        else:
            set_number.append(element)
    set_number.sort()
    incoming['all_doc'] = files*len(set_number)
    incoming['number_instance'] = set_number
    return {'error': False, 'data': incoming}


def check_create_account_number(incoming: dict) -> dict:
    start_path = incoming['start_path']
    if not start_path:
        return {'error': True, 'data': 'Путь к файлу учетных номеров пуст'}
    if not os.path.isdir(start_path):
        return {'error': True, 'data': 'Указанный путь к файлу учетных номеров не является директорией'}
    incoming['number_start']  = " ".join(incoming['account_number'].split())
    if not incoming['number_start'] :
        return {'error': True, 'data': 'Нет начального учетного номера'}
    for el in incoming['number_start']:
        if not re.match(r'[A-Za-z0-9\s]', el):
            return {'error': True, 'data': 'Некорректные символы в учетном номере'}
    return {'error': False, 'data': incoming}


def check_print_files(incoming: dict) -> dict:
    start_path = incoming['start_path']
    if not start_path:
        return {'error': True, 'data': 'Путь к файлам для печати пуст'}
    if not os.path.isdir(start_path):
        return {'error': True, 'data': 'Указанный путь к файлам для печати не является директорией'}
    incoming['all_doc'] = len(os.listdir(start_path))
    return {'error': False, 'data': incoming}


def check_print_certification(incoming: dict) -> dict:
    start_path = incoming['start_path']
    if not start_path:
        return {'error': True, 'data': 'Путь к файлам для печати пуст'}
    if incoming['duplex'] is False and incoming['last_duplex'] is False and incoming['one_side'] is False:
        return {'error': True, 'data': 'Не указан метод печати'}
    if not incoming['name_printer']:
        return {'error': True, 'data': 'Не выбран принтер'}
    if not os.path.isdir(start_path) and not os.path.isfile(start_path):
        return {'error': True, 'data': 'Указанная папка или файл удалена или переименована'}
    incoming['all_doc'] = len(os.listdir(start_path)) if os.path.isdir(start_path) else 1
    return {'error': False, 'data': incoming}


def check_word2pdf(incoming: dict) -> dict:
    start_path = incoming['start_path']
    if not start_path:
        return {'error': True, 'data': 'Путь к файлам для преобразования Word в PDF пуст'}
    if not os.path.isdir(start_path):
        return {'error': True, 'data': 'Путь к файлам для преобразования Word в PDF удалён или переименован'}
    finish_path = incoming['finish_path']
    if not finish_path:
        return {'error': True, 'data': 'Путь к конечной папке для преобразования Word в PDF пуст'}
    if not os.path.isdir(finish_path):
        return {'error': True, 'data': 'Путь к конечной папке для преобразования Word в PDF удален или переименован'}
    incoming['all_doc'] = 0
    for file in Path(start_path).rglob('*.*'):
        if file.suffix != '.docx':
            continue
        incoming['all_doc'] += 1
    if incoming['all_doc'] == 0:
        return {'error': True, 'data': 'В указанной папке для преобразования Word в PDF'
                                       ' нет подходящих файлов для преобразования'}
    return {'error': False, 'data': incoming}
