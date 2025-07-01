import os
import re
import shutil
import traceback

import fitz
import pythoncom
import docx
from pathlib import Path
from docx.shared import Pt, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENTATION

from small_functions import pages_count, delete_header_footer_second_acc
from word2pdf import word2pdf


def insert_header(doc, text_first_header, text_for_foot, fso_, text_finish: str = ''):
    header_1 = doc.sections[0].first_page_header  # Верхний колонтитул первой страницы
    head_1 = header_1.paragraphs[0]  # Параграф
    head_1.insert_paragraph_before(text_first_header)  # Вставляем перед колонтитулом
    head_1 = header_1.paragraphs[0]  # Выбираем новый первый параграф
    for header_styles in head_1.runs:
        header_styles.font.size = Pt(11)
        header_styles.font.name = 'Times New Roman'
    head_1_format = head_1.paragraph_format  # Настройки параграфа
    head_1_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT  # Выравниваем по правому краю
    footer_ = doc.sections[0].first_page_footer  # Нижний колонтитул первой страницы
    foot_ = footer_.paragraphs[0]  # Параграф
    foot_.text = text_for_foot  # Текст
    for foot_run in foot_.runs:
        foot_run.font.size = Pt(11)
        foot_run.font.name = 'Times New Roman'
    foot_format_ = foot_.paragraph_format  # Настройки параграфа
    foot_format_.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT  # Выравнивание по левому краю
    doc.sections[0].footer.paragraphs[0].text = text_for_foot  # Номера для страниц
    # Выравниваем по левому краю
    doc.sections[0].footer.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    doc.add_section()  # Добавляем последнюю страницу
    last_ = doc.sections[len(doc.sections) - 1].first_page_header  # Колонтитул для последней страницы
    last_.is_linked_to_previous = False  # Отвязываем от предыдущей секции чтобы не повторялись
    foot_ = doc.sections[len(doc.sections) - 1].first_page_footer  # Нижний колонтитул
    foot_.is_linked_to_previous = False  # Отвязываем
    # Текст для фонарика
    # if len(text_finish) > 0:
    foot_.paragraphs[0].text = text_finish
    # else:
    #     foot_.paragraphs[0].text = "Уч. № " + text_for_foot + \
    #                                "\nОтп. 1 экз. в адрес\n" + hdd_number + \
    #                                "\nИсп. " + executor + "\nПеч. " + print_people + "\n" + \
    #                                date + "\nб/ч"
    for footer_style in foot_.paragraphs[0].runs:
        footer_style.font.size = Pt(11)
        footer_style.font.name = 'Times New Roman'
    # if fso_:
    #     if 'заключение' in name_file_.lower() or 'акт' in name_file_.lower():
    #         path_new = path_new + '\\' + 'Материалы по специальной проверке технических средств'
    #     else:
    #         path_new = path_new + '\\' + 'Материалы по специальным исследованиям технических средств'
    #     try:
    #         os.mkdir(path_new)
    #     except FileExistsError:
    #         pass
    # doc.save(save_path)  # Сохраняем


def cell_write(style_for_doc, text_for_insert, table, number_rows=0):  # Заполнение ячеек в таблице в описи
    cells = table.rows[number_rows].cells  # Номер строки
    number_col = 0  # Номер столбца
    for elem in text_for_insert:
        cells[number_col].text = elem  # Заполняем элемент
        cells[number_col].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # Выравнивание по центру
        cells[number_col].paragraphs[0].style = style_for_doc
        if number_col == 1:  # Размер если ячейка с именем документа
            cells[number_col].width = 12801600  # 1.4 * 914400
        elif number_col == 3:  # Размер если ячейка с номером и грифом
            cells[number_col].width = 10972800  # 1.2 * 914400
        number_col += 1


def change_text(paragraphs, pattern: str, text: str, pt: int, string_date: bool = False) -> None:
    ru_month = {
        1: 'января',
        2: 'февраля',
        3: 'марта',
        4: 'апреля',
        5: 'мая',
        6: 'июня',
        7: 'июля',
        8: 'августа',
        9: 'сентября',
        10: 'октября',
        11: 'ноября',
        12: 'декабря'
    }
    if string_date:  # Если дату нужно заменить на строковое представление
        text = f"«{text.partition('.')[0]}»" \
               f" {ru_month[int(text.partition('.')[2].partition('.')[0])]}" \
               f" {text.rpartition('.')[2]} г."
    for paragraph in paragraphs:
        if re.findall(pattern, paragraph.text):
            insert_text = re.sub(pattern, text, paragraph.text)
            paragraph.text = insert_text
            for runs_ in paragraph.runs:
                runs_.font.size = Pt(pt)
                runs_.font.name = 'Times New Roman'
                runs_.font.bold = False
            break


def create_file(documents, data, pt_num, incoming_data, account_docs) -> dict:
    try:
        # Вероятно можно будет вынести создание колонтитула, изменение даты и сохранение документов из if
        para = False
        service = True if incoming_data['radioButton_FSB'] else False
        errors = []
        if data.action == 'copy':
            shutil.copy(str(Path(data.start_path)), str(Path(data.finish_path)))
            return {'status': 'success', 'text': f'Документ {data.name} скопирован', 'documents': documents}
        pythoncom.CoInitializeEx(0)
        if re.findall(r'опись', data.name.lower()):
            document = docx.Document()
        else:
            document = docx.Document(data.start_path)  # Открываем
        if re.findall(r'форма 3', data.name.lower()):
            pass
        if re.findall(r'приложение а', data.name.lower()):
            change_text(document.paragraphs, 'date', data.text, pt_num)
            change_text(document.paragraphs, r'date', data.date, pt_num, True)
        if re.findall(r'приложение', data.name.lower()):
            if re.findall(r'заключени', data.name.lower()):
                pass
                # if len(conclusion_num) == 0:
                #     if conclusion_number:
                #         conclusion_num_text = f'от {conclusion_number_date} № {str(conclusion_number)}'
                #     else:
                #         conclusion_num_text = False
                # elif len(conclusion_num) == 1:
                #     conclusion_num_text = f'от {date} № {str(conclusion_num[list(conclusion_num.keys())[0]])}'
                # else:
                #     # Такого случая не предусмотрено, если произошло - косяк.
                #     self.logging.warning('НЕСТАНДАРТНАЯ СИТУАЦИЯ, АЛГОРИТМ НЕ ПРОДУМАН И НЕ ОТЛАЖЕН')
                #     errors.append(f'В {name_el} не добавлен секретный номер, ситуация не согласована')
                #     conclusion_num_text = f'от {date} № '
                # if conclusion_num_text:
                #     for val_p, p in enumerate(doc.paragraphs):
                #         if re.findall(r'\[ЗАКЛНОМ]', p.text):
                #             text = re.sub(r'\[ЗАКЛНОМ]', conclusion_num_text, p.text)
                #             p.text = text
                #             for run in p.runs:
                #                 run.font.size = Pt(12)
                #                 run.font.name = 'Times New Roman'
                #             break
            else:
                change_text(document.paragraphs, r'\[АКТНОМ]', data.text, 12)
            change_text(document.paragraphs, r'date', data.date, 12)
        if re.findall(r'заключение', data.name.lower()):
            change_text(document.paragraphs, r'date', data.date, 12, True)
        if re.findall(r'протокол', data.name.lower()):
            change_text(document.paragraphs, r'\[ЗАКЛНОМ]', data.text_conclusion, pt_num)
            change_text(document.paragraphs, r'date', data.date, pt_num, True)
        if re.findall(r'предписание', data.name.lower()):
            change_text(document.paragraphs, r'\[ЗАКЛНОМ]', data.text_conclusion, pt_num)
            change_text(document.paragraphs, r'\[ПРОТНОМ]', data.text_protocol, pt_num)
            change_text(document.paragraphs, r'date', data.date, pt_num, True)
        if re.findall(r'акт', data.name.lower()):
            change_text(document.paragraphs, r'date', data.date, pt_num, True)
        if re.findall(r'утверждение', data.name.lower()):
            change_text(document.paragraphs, r'date', data.date, pt_num, True)
        if re.findall(r'опись', data.name.lower()):
            style = document.styles['Normal']
            font = style.font
            font.name = 'TimesNewRoman'
            font.size = Pt(12)
            section = document.sections[0]
            # section.orientation, section.page_width, section.page_height
            new_width, new_height = section.page_height, section.page_width  # Новые размеры
            section.orientation = WD_ORIENTATION.LANDSCAPE  # Альбомная ориентация
            section.page_width = new_width
            section.page_height = new_height
            section.left_margin = Cm(1.27)
            section.right_margin = Cm(1.27)
            section.top_margin = Cm(1.27)
            section.bottom_margin = Cm(1.27)
            section.different_first_page_header_footer = True
            # Добавляем необходимые надписи перед таблицей, выравниваем, создаем таблицу
            p = document.add_paragraph(f'Опись документов № {data.number}')
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            table = document.add_table(rows=1, cols=5, style='Table Grid')
            style = document.styles['Normal']
            font = style.font
            font.name = 'TimesNewRoman'
            font.size = Pt(12)
            cell_write(document.styles['Normal'], ['Порядковый номер', 'Наименование документа',
                                                   'Регистрационный номер', 'Номер экземпляра, гриф секретности',
                                                   'Количество листов в экземпляре'], table)
            for index, element in enumerate(account_docs.itertuples()):
                table.add_row()  # Добавляем колонку и значения
                style = document.styles['Normal']
                font = style.font
                font.name = 'TimesNewRoman'
                font.size = Pt(12)
                # сюда дописать все поля, подумать с надписью для описей - поместить в текст, там свободно
                cell_write(document.styles['Normal'],
                           [str(index + 1), *element.account_list_text.split('!')], table, index + 1)

            # Текст внизу таблицы
            p = document.add_paragraph()
            p.text = data.text
            p.paragraph_format.widow_control = True  # Чтобы подпись не убегала одна
            p.paragraph_format.keep_together = True  # Чтобы подпись не убегала одна
            # теперь подумать над новыми несекретными колонтитулами
        if re.findall(r'сопроводит', data.name.lower()):
            para = True if document.sections[0].different_first_page_header_footer else False
            for p in document.paragraphs:  # Для каждого параграфа
                if re.findall(r'registration_number', p.text):  # Ищем метку
                    p.text = re.sub(r'registration_number', f'{data.footer_text} от {data.date}', p.text)
                    for run in p.runs:
                        run.font.size = Pt(14)
                elif re.findall(r'Приложение:', p.text):
                    if account_docs.empty:
                        numbering = 1
                        if service:
                            ness_df = documents.loc[documents['name'].str.contains('заключение|предписание',
                                                                                   case=False)]
                        else:
                            ness_df = documents.loc[documents['name'].str.contains('заключение|протокол|предписание',
                                                                                   case=False)]
                        for file in ness_df.itertuples():
                            number_page = file.pages
                            page = 'листе' if int(number_page) == 1 else 'листах'
                            text = ''
                            if 'протокол' in file.name.lower():
                                account_doc = documents.loc[documents['name'].str.contains('приложение а', case=False) &
                                                            documents['number'].str.contains(file.number, case=False)]
                                if account_doc.empty is False:
                                    account_doc = account_doc.reset_index(drop=True)
                                    number_page = str(int(number_page) + int(account_doc.loc[0, 'pages']))
                                    page = 'листе' if int(number_page) == 1 else 'листах'
                                    page_app = 'листа' if int(account_doc.loc[0, 'pages']) > 1 else 'лист'
                                    text = f"{file.name.partition(' ')[0]}, уч. № {file.footer_text}," \
                                           f" экз.{file.num_scroll}, на {number_page} {page}, секретно," \
                                           f" {account_doc.loc[0, 'pages']} {page_app} - несекретно, только в адрес."
                            if len(text) == 0:
                                text = f"{file.name.partition(' ')[0]}, уч. № {file.footer_text}," \
                                       f" экз.{file.num_scroll}, на {number_page} {page}, секретно, только в адрес."
                            p.add_run('\n' + str(numbering) + '. ' + text)
                            p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT  # Выравниваем по левому краю
                            numbering += 1
                            for run in p.runs:
                                run.font.size = Pt(14)
                                run.font.name = 'Times New Roman'
                    else:
                        ness_df = documents.loc[documents['name'].str.contains('опись', case=False)]
                        for index, file in enumerate(ness_df.itertuples()):
                            number = file.number
                            pages = file.pages
                            page = 'листе' if pages == 1 else 'листах'  # Для правильной формулировки
                            footer_text = file.footer_text
                            text = f"Приложение согласно описи №{number} на {pages} {page}, уч. № {footer_text}," \
                                   f" экз. № 1, секретно, только в адрес."
                            p.add_run('\n' + str(index + 1) + '. ' + text)
                            p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT  # Выравниваем по левому краю
                            for run in p.runs:
                                run.font.size = Pt(14)
                                run.font.name = 'Times New Roman'
            if para:
                last = document.sections[len(document.sections) - 1].first_page_header  # Колонтитул последней страницы
                last.is_linked_to_previous = False  # Отвязываем от предыдущей секции чтобы не повторялись
                foot = document.sections[len(document.sections) - 1].first_page_footer  # Нижний колонтитул
                foot.is_linked_to_previous = False  # Отвязываем
            else:
                last = document.sections[len(document.sections) - 1].header  # Колонтитул для последней страницы
                last.is_linked_to_previous = False  # Отвязываем от предыдущей секции чтобы не повторялись
                foot = document.sections[len(document.sections) - 1].footer  # Нижний колонтитул
                foot.is_linked_to_previous = False  # Отвязываем
        if len(re.findall(r'приложение а', data.name.lower())) == 0:
            insert_header(document, data.first_header_text, data.footer_text, 'fso', data.text_finish)
        if Path(data.finish_path.parent).exists() is False:
            Path(data.finish_path.parent).mkdir(parents=True, exist_ok=True)
        document.save(data.finish_path)  # Сохраняем
        if re.findall(r'сопроводит', data.name.lower()) and re.findall(r'2 экз', data.name.lower()):
            answer = delete_header_footer_second_acc(data.finish_path, data.first_header_text, data.footer_text,
                                                     data.text_finish, para)
            if answer['status'] == 'error':
                errors.append(answer['text'])
        if incoming_data['checkBox_signature']:
            pattern = re.compile(r'\{\{\s[A-z]*\s}}')
            find_name = [pattern.findall(p.text)[0] for p in document.paragraphs if pattern.findall(p.text)]
            if len(find_name):
                pdf_path = Path(data.finish_path.parent, data.finish_path.name + '_copy1.pdf')
                word2pdf(str(data.finish_path), str(pdf_path))
                doc = fitz.open(str(pdf_path))
                page = doc.load_page(0)
                text_instances = {name[3: len(name) - 3]: page.search_for(name) for name in find_name}
                last_page = doc.load_page(len(doc) - 2)
                last_text_instances = {name[3: len(name) - 3]: last_page.search_for(name) for name in find_name}
                doc.close()
                os.remove(pdf_path)
                word2pdf(str(data.finish_path), str(pdf_path))
                doc = fitz.open(str(pdf_path))
                page = doc.load_page(0)
                last_page = doc.load_page(len(doc) - 2)
                for inst in text_instances:
                    for i in text_instances[inst]:
                        rect = fitz.Rect(i.x0 - 60, i.y0 - 60, i.x1 + 60, i.y1 + 60)
                        page.insert_image(rect, filename=str(incoming_data['signature_files'][inst]))
                for inst in last_text_instances:
                    for i in last_text_instances[inst]:
                        rect = fitz.Rect(i.x0 - 25, i.y0 - 25, i.x1 + 25, i.y1 + 25)
                        last_page.insert_image(rect, filename=str(incoming_data['signature_files'][inst]))
                doc.save(str(Path(data.finish_path.parent, data.finish_path.stem + '.pdf')))
                doc.close()
                os.remove(pdf_path)
        if re.findall(r'сопроводит', data.name.lower()) or re.findall(r'опись', data.name.lower()):
            pages = pages_count(data.finish_path)
            page = pages['pages'] - 1
            if page == 0:
                errors.append(f"Для файла {data.name} подсчёт кол-ва страниц завершился с ошибкой: {pages['text']}")
            index_doc = documents.loc[documents['name'] == data.name].index[0]
            documents.loc[index_doc, 'pages'] = page
        if errors:
            return {'status': 'warning', 'text': errors, 'documents': documents}
        return {'status': 'success', 'text': f'Документ {data.name} заполнен и сохранён', 'documents': documents}
        # не забыть посмотреть запрос
    except BaseException as ex:
        return {'status': 'error', 'text': f'Ошибка при создании документа {data.name}: {ex}',
                'trace': traceback.format_exc(),
                'documents': documents}
