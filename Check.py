import os
import re
import time

from PyQt5.QtWidgets import QMessageBox


def msgBox(title, text):
    msg = QMessageBox(QMessageBox.Critical, title, text)
    msg.exec_()


def doc_format(lineEdit_old, lineEdit_new, lineEdit_file_num, radioButton_FSB_df, radioButton_FSTEK_df,
               comboBox_classified, lineEdit_num_scroll, lineEdit_list_item, lineEdit_number, lineEdit_executor,
               lineEdit_conclusion, lineEdit_prescription, lineEdit_print, lineEdit_executor_acc_sheet, label_executor,
               label_conclusion, label_prescription, label_print, label_executor_acc_sheet, lineEdit_date,
               groupBox_inventory_insert, radioButton_40_num, radioButton_all_doc, lineEdit_account_post,
               lineEdit_account_signature, lineEdit_account_path, hdd_number, groupBox_form27_insert, lineEdit_firm,
               lineEdit_path_form_27_create, qroupBox_second_copy, checkBox_conclusion, checkBox_protocol,
               checkBox_preciption, q, log, package):  # Ф-я для проверки введенных значений
    def chek(n, e):  # Ф-я для проверки элементов
        f = 0
        for elem in e:
            if n == elem:
                f = 1
                return f
        return f

    package_ = True if package.isChecked() else False
    # Путь к исходным документам и проверки
    path_old = lineEdit_old.text().strip()
    if not path_old:
        msgBox('УПС!', 'Путь к исходным документам пуст')
        return
    if os.path.isdir(path_old):
        if len(os.listdir(path_old)) == 0:
            msgBox('УПС!', 'Папка с исходными документами пуста')
            return
        if package_:
            docs = []
            error = [i for i in os.listdir(path_old) if os.path.isfile(path_old + '\\' + i)]
            if error:
                msgBox('УПС!', 'В директории для пакетного преобразования присутствуют файлы')
                return
            for folder in os.listdir(path_old):
                # Ошибка если есть файлы старого формата
                docs = docs + [i for i in os.listdir(path_old + '\\' + folder) if i[-3:] == 'doc']

        else:
            error = [i for i in os.listdir(path_old) if os.path.isdir(path_old + '\\' + i) and
                     ('материалы' not in i.lower())]
            if error:
                msgBox('УПС!', 'В директории для преобразования присутствуют папки')
                return
            docs = [i for i in os.listdir(path_old) if i[-3:] == 'doc']  # Ошибка если есть файлы старого формата
        if len(docs) != 0:
            text = 'Файлы старого формата:\n' + '\n'.join(docs)
            msgBox('УПС!', text)
            return
    else:
        msgBox('УПС!', 'Указанный путь к исходным документам не является директорией')
        return
    # Путь к конечным документам и проверки
    path_new = lineEdit_new.text().strip()
    if not path_new:
        msgBox('УПС!', 'Путь к конечной папке пуст')
        return
    if os.path.isdir(path_new):
        pass
    else:
        msgBox('УПС!', 'Указанный путь к конечной папке не является директорией')
        return
    if len(os.listdir(path_new)) != 0:
        msgBox('УПС!', 'Конечная папка не пуста, очистите директорию')
        return
    file_num = lineEdit_file_num.text().strip()
    if file_num:
        if os.path.isdir(file_num):
            msgBox('УПС!', 'Указанный путь к файлу номеров является директорией')
            return
        else:
            if os.path.exists(file_num):
                if file_num.endswith('.xlsx'):
                    pass
                else:
                    msgBox('УПС!', 'Файл номеров не формата .xlsx')
                    return
            else:
                msgBox('УПС!', 'Файл номеров удалён или переименван')
                return
    # Ведомство
    if radioButton_FSB_df.isChecked():
        service = True
    else:
        if radioButton_FSTEK_df.isChecked():
            service = False
        else:
            msgBox('УПС!', 'Не выбрано ведомство для вставки колонтитулов')
            return
    # Гриф секретности
    class_ = {'ДСП': 'Для служебного пользования', 'С': 'Секретно', 'СС': 'Совершенно секретно',
              'ОВ': 'Особой важности'}
    classified = comboBox_classified.currentText().strip()
    if not classified:
        msgBox('УПС!', 'Не выбрана категория секретности')
        return
    classified = class_[classified]
    # Номер экземпляра
    num_scroll = lineEdit_num_scroll.text().strip()
    if not num_scroll:
        msgBox('УПС!', 'Не выбран номер экземпляра')
        return
    # Пункт перечня
    list_item = lineEdit_list_item.text().strip()
    if not list_item:
        msgBox('УПС!', 'Не указан пункт перечня')
        return
    # Номер
    number = lineEdit_number.text().strip()
    if not file_num:
        err_f = ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0', '/', 'c', 'с', '-', 'Н', 'С', 'с')
        if number[-1] in ['С', 'с']:
            number = number.replace(number[-1], 'c')
        if not number:
            msgBox('УПС!', 'Не указан номер')
            return
        for i in number:
            flag = chek(i, err_f)
            if not flag:
                msgBox('УПС!', 'Есть лишние символы в номере')
                return
        if (re.match(r'\w+/\w+/\w+c', number) is None) and (re.match(r'НС-\w+c', number) is None):
            msgBox('УПС!', 'Секретный номер указан неверно')
            return
    # Исполнитель, заключение, предписание, протокол, печать
    executor = lineEdit_executor.text().strip()
    conclusion = lineEdit_conclusion.text().strip()
    prescription = lineEdit_prescription.text().strip()
    print_people = lineEdit_print.text().strip()
    executor_acc_sheet = lineEdit_executor_acc_sheet.text().strip()
    list_label = [label_executor, label_conclusion, label_prescription, label_print, label_executor_acc_sheet]
    i = 0
    for element in [executor, conclusion, prescription, print_people, executor_acc_sheet]:
        if not element:
            msgBox('УПС!', 'Не указан(а) ' + list_label[i].text())
            return
        i += 1
    # Дата
    date = lineEdit_date.text().strip()
    try:
        valid_date = time.strptime(date, '%d.%m.%Y')
    except ValueError:
        msgBox('УПС!', 'Формат даты указан неверно! (необходимый формат: dd.mm.yyyy)')
        return
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
            msgBox('УПС!', 'Не указано количество документов в описе')
            return
        account_post = lineEdit_account_post.text().strip()
        if not account_post:
            msgBox('УПС!', 'Не указана должность для описи')
            return
        account_signature = lineEdit_account_signature.text().strip()
        if not account_signature:
            msgBox('УПС!', 'Не указана подпись для описи')
            return
        account_path = lineEdit_account_path.text().strip()
        if not account_path:
            msgBox('УПС!', 'Не указан путь для файла описи')
            return
        else:
            if os.path.isdir(account_path):
                pass
            else:
                msgBox('УПС!', 'Для описи необходимо указать файл')
                return
    if not hdd_number:
        msgBox('УПС!', 'Отсутствует номер жесткого диска')
        return
    firm = None
    form_27 = None
    if groupBox_form27_insert.isChecked():
        firm = lineEdit_firm.text().strip()
        if not firm:
            msgBox('УПС!', 'Не заполнена организация для 27 формы')
            return
        form_27 = lineEdit_path_form_27_create.text().strip()
        if not form_27:
            msgBox('УПС!', 'Нет пути для 27 формы')
            return
        else:
            if os.path.isdir(form_27):
                pass
            else:
                msgBox('УПС!', 'Указанный путь для 27 формы не является директорией')
                return
    second_copy = []
    if qroupBox_second_copy.isChecked():
        second_copy = [True if i.isChecked() else False for i in [checkBox_conclusion, checkBox_protocol,
                                                                  checkBox_preciption]]
        if all(i is False for i in second_copy):
            msgBox('УПС!', 'Не выбран ни один документ для второго экземпляра')
            return

    return [path_old, path_new, file_num, classified, num_scroll, list_item,
            number, executor, conclusion, prescription, print_people, date, executor_acc_sheet, account,
            flag_inventory, account_post, account_signature, account_path, firm, form_27, second_copy, service,
            hdd_number, q, log, package_]


def doc_print(radioButton_FSB_print, radioButton_FSTEK_print, lineEdit_old_print, lineEdit_account_numbers,
              checkBox_add_account_numbers, lineEdit_add_account_numbers, checkBox_form_27, lineEdit_path_form_27_print,
              button_gr, lineEdit_printer, checkBox_print_order, path_for_default, log, package):
    # Ведомство
    package_ = True if package.isChecked() else False
    if radioButton_FSB_print.isChecked():
        service = True
    else:
        if radioButton_FSTEK_print.isChecked():
            service = False
        else:
            msgBox('УПС!', 'Не выбрано ведомство при печати документов')
            return
    path_old_print = lineEdit_old_print.text().strip()
    if not path_old_print:
        msgBox('УПС!', 'Путь к исходным документам для печати пуст')
        return
    if os.path.isdir(path_old_print):
        if len(os.listdir(path_old_print)) == 0:
            msgBox('УПС!', 'Папка с исходными документами для печати пуста')
            return
        docs = []
        if package_:
            error = [i for i in os.listdir(path_old_print) if os.path.isfile(path_old_print + '\\' + i)]
            print(error)
            if error:
                msgBox('УПС!', 'В директории для пакетной печати присутствуют файлы')
                return
            for folder in os.listdir(path_old_print):
                # Ошибка если есть файлы старого формата
                docs = docs + [i for i in os.listdir(path_old_print + '\\' + folder) if i[-3:] == 'doc']
        else:
            error = [i for i in os.listdir(path_old_print) if os.path.isdir(path_old_print + '\\' + i)]
            if error:
                msgBox('УПС!', 'В директории для преобразования присутствуют папки')
                return
            docs = [i for i in os.listdir(path_old_print) if i[-3:] == 'doc']  # Ошибка если есть файлы старого формата
        if len(docs) != 0:
            text = 'Файлы старого формата:\n' + '\n'.join(docs)
            msgBox('УПС!', text)
            return
    else:
        msgBox('УПС!', 'Указанный путь к исходным документам для печати не является директорией')
        return
    # Путь к номерам
    path_account_num = lineEdit_account_numbers.text().strip()
    add_path_account_num = None
    if checkBox_add_account_numbers.isChecked():
        add_path_account_num = lineEdit_add_account_numbers.text().strip()
        if not add_path_account_num:
            msgBox('УПС!', 'Путь к доп. файлу номеров учетных листов пуст')
            return
        if os.path.isdir(add_path_account_num):
            msgBox('УПС!', 'Указанный путь к доп. файлу номеров учётных листов является директорией')
            return
        else:
            if os.path.exists(add_path_account_num):
                if add_path_account_num.endswith('.xlsx'):
                    pass
                else:
                    msgBox('УПС!', 'Доп. файл номеров не формата .xlsx')
            else:
                msgBox('УПС!', 'Доп. файл номеров удалён или переименван')
                return
    if not path_account_num:
        msgBox('УПС!', 'Путь к файлу номеров учетных листов пуст')
        return
    if os.path.isdir(path_account_num):
        msgBox('УПС!', 'Указанный путь к файлу номеров учётных листов является директорией')
        return
    else:
        if os.path.exists(path_account_num):
            if path_account_num.endswith('.xlsx'):
                pass
            else:
                msgBox('УПС!', 'Файл номеров не формата .xlsx')
        else:
            msgBox('УПС!', 'Файл номеров удалён или переименван')
            return
    # Форма 27
    path_form_27 = False
    if checkBox_form_27.isChecked():
        if package_ is False:
            path_form_27 = lineEdit_path_form_27_print.text().strip()
            if not path_form_27:
                msgBox('УПС!', 'Путь к 27 форме пуст')
                return
            if os.path.isdir(path_form_27):
                msgBox('УПС!', 'Указанный путь к 27 форме является директорией')
                return
            else:
                if os.path.exists(path_form_27) and not package_:
                    if path_form_27.endswith('.xlsx'):
                        pass
                    else:
                        msgBox('УПС!', 'Файл "Форма 27" не формата .xlsx')
                else:
                    if not package_:
                        msgBox('УПС!', 'Файл "Форма 27" удалён или переименван')
                        return
        else:
            path_form_27 = True
    # Способ печати
    print_flag = None
    for button in button_gr:
        if button.isChecked():
            print_flag = button.text()
    if not print_flag:
        msgBox('УПС!', 'Не указан метод печати')
        return
    print_name = lineEdit_printer.text().strip()
    if not print_name:
        msgBox('УПС!', 'Не выбран принтер')
        return
    print_order = True if checkBox_print_order.isChecked() else False
    return [path_old_print, path_account_num, add_path_account_num, print_flag,
            print_name, path_form_27, print_order, service, path_for_default, log, package_]
