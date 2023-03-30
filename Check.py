import os
import re
import time


def doc_format(lineEdit_old, lineEdit_new, lineEdit_file_num, radioButton_FSB_df, radioButton_FSTEK_df,
               comboBox_classified, lineEdit_num_scroll, lineEdit_list_item, lineEdit_number, lineEdit_protocol,
               lineEdit_conclusion, lineEdit_prescription, lineEdit_print, lineEdit_executor_acc_sheet, label_protocol,
               label_conclusion, label_prescription, label_print, label_executor_acc_sheet, lineEdit_date, lineEdit_act,
               lineEdit_statement,
               groupBox_inventory_insert, radioButton_40_num, radioButton_all_doc, lineEdit_account_post,
               lineEdit_account_signature, lineEdit_account_path, hdd_number, groupBox_form27_insert, lineEdit_firm,
               lineEdit_path_form_27_create, qroupBox_instance, lineEdit_number_instance, checkBox_conclusion,
               checkBox_protocol, checkBox_preciption, package, action_MO):  # Ф-я для проверки введенных значений

    def check(n, e):
        for symbol in e:
            if n == symbol:
                return False
        return True

    package_ = True if package.isChecked() else False
    action_MO_ = True if action_MO.isChecked() else False
    # Путь к исходным документам и проверки
    path_old = lineEdit_old.text().strip()
    if not path_old:
        return ['УПС!', 'Путь к исходным документам пуст']
    if os.path.isdir(path_old):
        if len(os.listdir(path_old)) == 0:
            return ['УПС!', 'Папка с исходными документами пуста']
        if package_:
            docs = []
            for folder in os.listdir(path_old):
                if os.path.isdir(path_old + '\\' + folder):
                    # Ошибка если есть файлы старого формата
                    docs += [i for i in os.listdir(path_old + '\\' + folder) if i[-3:] == 'doc']

        else:
            error = [i for i in os.listdir(path_old) if os.path.isdir(path_old + '\\' + i) and
                     ('материалы' not in i.lower())]
            if error:
                return ['УПС!', 'В директории для преобразования присутствуют папки']
            docs = [i for i in os.listdir(path_old) if i[-3:] == 'doc']  # Ошибка если есть файлы старого формата
        if len(docs) != 0:
            text = 'Файлы старого формата:\n' + '\n'.join(docs)
            return ['УПС!', text]
    else:
        return ['УПС!', 'Указанный путь к исходным документам не является директорией']
    if action_MO_:
        file_mo = [mo for mo in os.listdir(path_old) if mo[-3:] == 'txt' and 'F19' in mo]
        if not file_mo:
            return ['УПС!', 'Нет текстового файла с серийниками для создания отчёта для МО']
    # Путь к конечным документам и проверки
    path_new = lineEdit_new.text().strip()
    if not path_new:
        return ['УПС!', 'Путь к конечной папке пуст']
    if os.path.isdir(path_new):
        pass
    else:
        return ['УПС!', 'Указанный путь к конечной папке не является директорией']
    if len(os.listdir(path_new)) != 0:
        return ['УПС!', 'Конечная папка не пуста, очистите директорию']
    file_num = lineEdit_file_num.text().strip()
    if file_num:
        if os.path.isdir(file_num):
            return ['УПС!', 'Указанный путь к файлу номеров является директорией']
        else:
            if os.path.exists(file_num):
                if file_num.endswith('.xlsx'):
                    pass
                else:
                    return ['УПС!', 'Файл номеров не формата .xlsx']
            else:
                return ['УПС!', 'Файл номеров удалён или переименван']
    # Ведомство
    if radioButton_FSB_df.isChecked():
        service = True
    elif radioButton_FSTEK_df.isChecked():
        service = False
    else:
        return ['УПС!', 'Не выбрано ведомство для вставки колонтитулов']
    # Гриф секретности
    class_ = {'ДСП': 'Для служебного пользования', 'С': 'Секретно', 'СС': 'Совершенно секретно',
              'ОВ': 'Особой важности'}
    classified = comboBox_classified.currentText().strip()
    if not classified:
        return ['УПС!', 'Не выбрана категория секретности']
    classified = class_[classified]
    # Номер экземпляра
    num_scroll = lineEdit_num_scroll.text().strip()
    if not num_scroll:
        return ['УПС!', 'Не выбран номер экземпляра']
    # Пункт перечня
    list_item = lineEdit_list_item.text().strip()
    if not list_item:
        return ['УПС!', 'Не указан пункт перечня']
    # Номер
    number = lineEdit_number.text().strip()
    if not file_num:
        if number[-1] in ['С', 'с']:
            number = number.replace(number[-1], 'c')
        if not number:
            return ['УПС!', 'Не указан номер']
        for i in number:
            if check(i, ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0', '/', 'c', 'с', '-', 'Н', 'С', 'с')):
                return ['УПС!', 'Есть лишние символы в номере']
        if (re.match(r'\w+/\w+/\w+c$', number) is None) and (re.match(r'НС-\w+c$', number) is None):
            return ['УПС!', 'Секретный номер указан неверно']
    # Исполнитель, заключение, предписание, протокол, печать
    act = lineEdit_act.text().strip()
    statement = lineEdit_statement.text().strip()
    protocol = lineEdit_protocol.text().strip()
    conclusion = lineEdit_conclusion.text().strip()
    prescription = lineEdit_prescription.text().strip()
    print_people = lineEdit_print.text().strip()
    executor_acc_sheet = lineEdit_executor_acc_sheet.text().strip()
    list_label = [label_protocol, label_conclusion, label_prescription, label_print, label_executor_acc_sheet]
    i = 0
    for element in [protocol, conclusion, prescription, print_people, executor_acc_sheet]:
        if not element:
            return ['УПС!', 'Не указан(а) ' + list_label[i].text()]
        i += 1
    # Дата
    date = lineEdit_date.text().strip()
    try:
        time.strptime(date, '%d.%m.%Y')
    except ValueError:
        return ['УПС!', 'Формат даты указан неверно! (необходимый формат: dd.mm.yyyy)']
    account = None
    flag_inventory = None
    account_post = None
    account_signature = None
    account_path = None
    if groupBox_inventory_insert.isChecked() and package_ is False:
        account = True
        if radioButton_40_num.isChecked() or radioButton_all_doc.isChecked():
            flag_inventory = 40 if radioButton_40_num.isChecked() else 1
        else:
            return ['УПС!', 'Не указано количество документов в описе']
        account_post = lineEdit_account_post.text().strip()
        if not account_post:
            return ['УПС!', 'Не указана должность для описи']
        account_signature = lineEdit_account_signature.text().strip()
        if not account_signature:
            return ['УПС!', 'Не указана подпись для описи']
        account_path = lineEdit_account_path.text().strip()
        if not account_path:
            return ['УПС!', 'Не указан путь для файла описи']
        else:
            if os.path.isdir(account_path):
                pass
            else:
                return ['УПС!', 'Для описи необходимо указать файл']
    if not hdd_number:
        return ['УПС!', 'Отсутствует номер жесткого диска']
    firm = None
    form_27 = None
    if groupBox_form27_insert.isChecked():
        firm = lineEdit_firm.text().strip()
        if not firm:
            return ['УПС!', 'Не заполнена организация для 27 формы']
        form_27 = lineEdit_path_form_27_create.text().strip()
        if not form_27:
            return ['УПС!', 'Нет пути для 27 формы']
        else:
            if os.path.isdir(form_27):
                pass
            else:
                return ['УПС!', 'Указанный путь для 27 формы не является директорией']
    second_copy = []
    complect = None
    if qroupBox_instance.isChecked():
        number_instance = lineEdit_number_instance.text().strip()
        for i in number_instance:
            if check(i, ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0', ' ', '-', ',', '.')):
                return ['УПС!', 'Есть лишние символы в номерах экземпляров']
        complect_num = number_instance.replace(' ', '').replace(',', '.')
        if complect_num[0] == '.' or complect_num[0] == '-':
            return ['УПС!', 'Первый символ введён не верно']
        if complect_num[-1] == '.' or complect_num[-1] == '-':
            return ['УПС!', 'Последний символ введён не верно']
        for i in range(len(complect_num)):
            if complect_num[i] == '.' or complect_num[i] == '-':
                if complect_num[i + 1] == '.' or complect_num[i + 1] == '-':
                    return ['УПС!', 'Два разделителя номеров подряд']
        complect = []
        for element in complect_num.split('.'):
            if '-' in element:
                num1, num2 = int(element.partition('-')[0]), int(element.partition('-')[2])
                if num1 >= num2:
                    return ['УПС!', 'Диапазон номеров экземпляров указан не верно']
                else:
                    for el in range(num1, num2 + 1):
                        complect.append(el)
            else:
                complect.append(element)
        complect.sort()
        if len(complect) != len(set(complect)):
            return ['УПС!', 'Есть повторения в номерах экземпляров']
        second_copy = [True if i.isChecked() else False for i in [checkBox_conclusion, checkBox_protocol,
                                                                  checkBox_preciption]]
        if all(i is False for i in second_copy):
            return ['УПС!', 'Не выбран ни один документ для второго экземпляра']
    return {'path_old': path_old, 'path_new': path_new, 'file_num': file_num, 'classified': classified,
            'num_scroll': num_scroll, 'list_item': list_item, 'number': number, 'protocol': protocol,
            'conclusion': conclusion, 'prescription': prescription, 'print_people': print_people, 'date': date,
            'executor_acc_sheet': executor_acc_sheet, 'account': account, 'flag_inventory': flag_inventory,
            'account_post': account_post, 'account_signature': account_signature, 'account_path': account_path,
            'firm': firm, 'form_27': form_27, 'second_copy': second_copy, 'service': service, 'hdd_number': hdd_number,
            'package': package_, 'action_MO': action_MO_, 'act': act, 'statement': statement,
            'number_instance': complect}


def doc_print(radioButton_FSB_print, radioButton_FSTEK_print, checkBox_conclusion_print, checkBox_protocol_print,
              checkBox_preciption_print, lineEdit_old_print, lineEdit_account_numbers,
              checkBox_add_account_numbers, lineEdit_add_account_numbers, checkBox_form_27, lineEdit_path_form_27_print,
              button_gr, lineEdit_printer, checkBox_print_order, path_for_default, package):
    # Ведомство
    package_ = True if package.isChecked() else False
    if radioButton_FSB_print.isChecked():
        service = True
    elif radioButton_FSTEK_print.isChecked():
        service = False
    else:
        return ['УПС!', 'Не выбрано ведомство при печати документов']
    document_list = {i: True if j.isChecked() else False for i, j in zip(['заключение', 'протокол', 'предписание'],
                                                                         [checkBox_conclusion_print,
                                                                          checkBox_protocol_print,
                                                                          checkBox_preciption_print])}
    path_old_print = lineEdit_old_print.text().strip()
    if not path_old_print:
        return ['УПС!', 'Путь к исходным документам для печати пуст']
    if os.path.isdir(path_old_print):
        if len(os.listdir(path_old_print)) == 0:
            return ['УПС!', 'Папка с исходными документами для печати пуста']
        docs = []
        if package_:
            error = [i for i in os.listdir(path_old_print) if os.path.isfile(path_old_print + '\\' + i)]
            print(error)
            if error:
                return ['УПС!', 'В директории для пакетной печати присутствуют файлы']
            for folder in os.listdir(path_old_print):
                # Ошибка если есть файлы старого формата
                docs = docs + [i for i in os.listdir(path_old_print + '\\' + folder) if i[-3:] == 'doc']
        else:
            error = [i for i in os.listdir(path_old_print) if os.path.isdir(path_old_print + '\\' + i)]
            if error:
                return ['УПС!', 'В директории для преобразования присутствуют папки']
            docs = [i for i in os.listdir(path_old_print) if i[-3:] == 'doc']  # Ошибка если есть файлы старого формата
        if len(docs) != 0:
            text = 'Файлы старого формата:\n' + '\n'.join(docs)
            return ['УПС!', text]
    else:
        return ['УПС!', 'Указанный путь к исходным документам для печати не является директорией']
    # Путь к номерам
    path_account_num = lineEdit_account_numbers.text().strip()
    add_path_account_num = None
    if checkBox_add_account_numbers.isChecked():
        add_path_account_num = lineEdit_add_account_numbers.text().strip()
        if not add_path_account_num:
            return ['УПС!', 'Путь к доп. файлу номеров учетных листов пуст']
        if os.path.isdir(add_path_account_num):
            return ['УПС!', 'Указанный путь к доп. файлу номеров учётных листов является директорией']
        else:
            if os.path.exists(add_path_account_num):
                if add_path_account_num.endswith('.xlsx'):
                    pass
                else:
                    return ['УПС!', 'Доп. файл номеров не формата .xlsx']
            else:
                return ['УПС!', 'Доп. файл номеров удалён или переименван']
    if not path_account_num:
        return ['УПС!', 'Путь к файлу номеров учетных листов пуст']
    if os.path.isdir(path_account_num):
        return ['УПС!', 'Указанный путь к файлу номеров учётных листов является директорией']
    else:
        if os.path.exists(path_account_num):
            if path_account_num.endswith('.xlsx'):
                pass
            else:
                return ['УПС!', 'Файл номеров не формата .xlsx']
        else:
            return ['УПС!', 'Файл номеров удалён или переименван']
    # Форма 27
    path_form_27 = False
    if checkBox_form_27.isChecked():
        if package_ is False:
            path_form_27 = lineEdit_path_form_27_print.text().strip()
            if not path_form_27:
                return ['УПС!', 'Путь к 27 форме пуст']
            if os.path.isdir(path_form_27):
                return ['УПС!', 'Указанный путь к 27 форме является директорией']
            else:
                if os.path.exists(path_form_27) and not package_:
                    if path_form_27.endswith('.xlsx'):
                        pass
                    else:
                        return ['УПС!', 'Файл "Форма 27" не формата .xlsx']
                else:
                    if not package_:
                        return ['УПС!', 'Файл "Форма 27" удалён или переименван']
        else:
            path_form_27 = True
    # Способ печати
    print_flag = None
    for button in button_gr:
        if button.isChecked():
            print_flag = button.text()
    if not print_flag:
        return ['УПС!', 'Не указан метод печати']
    print_name = lineEdit_printer.text().strip()
    if not print_name:
        return ['УПС!', 'Не выбран принтер']
    print_order = True if checkBox_print_order.isChecked() else False
    return {'path_old_print': path_old_print, 'path_account_num': path_account_num,
            'add_path_account_num': add_path_account_num, 'print_flag': print_flag, 'name_printer': print_name,
            'path_form_27': path_form_27, 'print_order': print_order, 'service': service,
            'path_for_default': path_for_default, 'package_': package_, 'document_list': document_list}
