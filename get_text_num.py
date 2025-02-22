import logging
import traceback
import re
from pathlib import Path
import pandas as pd
from small_functions import pages_count


def create_text_for_docs(log: logging, docs: list, documents: pd.DataFrame, num_1: str,	 num_2: str,
						 add_conclusion_num: str, add_conclusion_date: str, user_date:str, conclusion: str,
						 protocol: str, prescription: str, print: str, act: str, statement: str, classified: str,
						 num_scroll: str, list_item: str, hdd_number: str, dict_40: list, for_27: list, line_doing,
						 add_list_item: str = '', dict_file=None
						 ) -> dict:
	"""Вставка в основную таблицу номеров, текстовок, дат, исполнителей. Возвращает части секретного номера,
	чтобы продолжить, если включен пакетный режим"""
	if dict_file is None:
		dict_file = {}
	# Имена чтобы не увеличивался секретный номер
	doc_name_for_continue = ['result', 'инфокарта', 'приложение а']
	doc_name_for_dict_40 = ['акт', 'заключение', 'протокол', 'предписание', 'утверждение']
	try:
		errors = []
		for doc in docs:  # Для файлов в папке
			doc_name = doc.name
			parent_path = doc.parent
			line_doing.emit(f'Генерируем файл {doc_name}')
			logging.info(f"Добавляем данные для {doc_name}")
			index_doc = documents.loc[documents['start_path'] == doc].index[0]
			# Определяем есть ли номер. Сначала отсекаем, потом ищем структуру.
			number_doc = doc_name.rpartition('.')[0].rpartition(' ')[2]
			if re.findall(r'\d\.\d', number_doc) is False:
				number_doc = False
			text_first_header = classified + '\n' + list_item + '\nЭкз. №' + num_scroll
			footer_text = False
			date = user_date
			executor = False
			if dict_file and doc_name.rpartition('.')[0] in dict_file:  # Если есть файл номеров
				footer_text = dict_file[doc_name.rpartition('.')[0]][0]  # Текст для нижнего колонтитула
				date = dict_file[doc_name.rpartition('.')[0]][1]  # Дата
			else:
				if all([True if _ not in doc_name.lower() else False for _ in doc_name_for_continue]):
					footer_text = num_1 + num_2 + 'c'  # Текст для нижнего колонтитула
			if doc_name.lower() == 'форма 3.docx':
				executor = print
			elif 'result' in doc_name.lower():
				documents.loc[index_doc, 'action'] = 'copy'
				# documents[doc]['action'] = 'copy'
				continue
			elif re.findall('сопроводит', doc_name.lower()) or re.findall('запрос', doc_name.lower()):
				documents.loc[index_doc, 'action'] = 'acc_doc'
				# documents[doc]['action'] = 'acc_doc'
			elif re.findall('инфокарта', doc_name.lower()):
				documents.loc[index_doc, 'action'] = 'copy'
				documents.loc[index_doc, 'number'] = number_doc
				# documents[doc]['action'] = 'copy'
			# pythoncom.CoInitializeEx(0)
			# status.emit('Форматируем документ ' + name_el)
			# doc = docx.Document(el_)  # Открываем
			elif re.findall(r'приложение а', doc_name.lower()):
				# index_doc = documents.loc[documents['start_path'] == doc].index[0]
				# number_protocol = doc_name.rpartition(' ')[2].rpartition('.')[0]
				documents.loc[index_doc, 'number'] = number_doc
				protocol_name = documents[(documents['name'].str.contains(r'протокол', case=False))
									  & (documents['name'].str.contains(number_doc)
										 & (documents['parent_path'] == parent_path))]
				protocol_name = protocol_name.reset_index(drop=True)
				if len(protocol_name) > 0:
					documents.loc[index_doc, 'text'] = f"к протоколу уч. № {str(protocol_name.loc[0, 'footer_text'])} от date"
					documents.loc[index_doc, 'change_date'] = True
				else:
					errors.append(f"Для файла {doc_name} не нашли протокол, документ не заполнен")
			elif re.findall(r'приложение', doc_name.lower()):
				# index_doc = documents.loc[documents['start_path'] == doc].index[0]
				if re.findall(r'заключени', doc_name.lower()):
					conclusion = documents[(documents['name'].str.contains(r'заключение'))]
					if len(conclusion) == 0:
						documents.loc[index_doc, 'text'] = f"от {add_conclusion_date} № {str(add_conclusion_num)}"
					elif len(conclusion) == 1:
						documents.loc[index_doc, 'text'] = f"от {add_conclusion_date} № {str(add_conclusion_num)}"
					else:
						# Такого случая не предусмотрено, если произошло - косяк.
						log.warning('НЕСТАНДАРТНАЯ СИТУАЦИЯ, АЛГОРИТМ НЕ ПРОДУМАН И НЕ ОТЛАЖЕН')
						errors.append(f'В {doc_name} не добавлен секретный номер, ситуация не согласована')
						documents.loc[index_doc, 'text'] = f"от {date} № "
					executor = protocol
					# documents.loc[index_doc, 'executor'] = protocol
				else:
					act_number = documents[(documents['name'].str.contains(r'акт'))]
					documents.loc[index_doc, 'text'] = f"от date № {act_number}"
					executor = act
					# documents.loc[index_doc, 'executor'] = act
				documents.loc[index_doc, 'change_date'] = True
				if add_list_item:
					text_first_header = classified + '\n' + add_list_item + '\nЭкз. №' + num_scroll
			elif re.findall(r'заключение', doc_name.lower()):
				# number_conclusion = doc_name.rpartition('.')[0].rpartition(' ')[2]
				documents.loc[index_doc, 'number'] = number_doc
				executor = conclusion
				# documents.loc[index_doc, 'executor'] = conclusion
				documents.loc[index_doc, 'change_date'] = True
			elif re.findall(r'протокол', doc_name.lower()):
				# Параграф для колонтитула первой страницы
				if add_list_item:
					text_first_header = classified + '\n' + add_list_item + '\nЭкз. №' + num_scroll
				# number_protocol = doc_name.rpartition('.')[0].rpartition(' ')[2]
				documents.loc[index_doc, 'number'] = number_doc
				conclusion_name = documents[(documents['name'].str.contains(r'заключение', case=False)
										& (documents['parent_path'] == parent_path))]
				if len(conclusion_name) > 1:
					conclusion_name = documents[(documents['name'].str.contains(r'заключение', case=False)
											& (documents['number'].str.contains(number_doc))
											& (documents['parent_path'] == parent_path))]
				conclusion_name = conclusion_name.reset_index(drop=True)
				if len(conclusion_name) == 0:
					documents.loc[index_doc, 'text_conclusion'] = f"уч. № {str(add_conclusion_num)} от {add_conclusion_date}"
				else:
					documents.loc[index_doc, 'text_conclusion'] = f"уч. № {str(conclusion_name.loc[0, 'footer_text'])} от {conclusion_name.loc[0, 'date']}"
				# documents.loc[index_doc, 'executor'] = protocol
				executor = protocol
				documents.loc[index_doc, 'change_date'] = True
			elif re.findall(r'предписание', doc_name.lower()):
				# Параграф для колонтитула первой страницы
				if add_list_item:
					text_first_header = classified + '\n' + add_list_item + '\nЭкз. №' + num_scroll
				# name_prescription = doc_name.rpartition('.')[0].rpartition(' ')[0]
				# number_prescription = doc_name.rpartition('.')[0].rpartition(' ')[2]
				documents.loc[index_doc, 'number'] = number_doc
				conclusion_name = documents[(documents['name'].str.contains(r'заключение', case=False)
										& (documents['parent_path'] == parent_path))]
				if len(conclusion_name) > 1:
					conclusion = documents[(documents['name'].str.contains(r'заключение')
											& (documents['number'].str.contains(number_doc))
											& (documents['parent_path'] == parent_path))]
				protocol_name = documents[(documents['name'].str.contains(r'протокол', case=False)
										& (documents['number'].str.contains(number_doc))
										& (documents['parent_path'] == parent_path))]
				conclusion_name = conclusion_name.reset_index(drop=True)
				protocol_name = protocol_name.reset_index(drop=True)
				if len(protocol_name):
					documents.loc[index_doc, 'text_protocol'] = f"уч. № {str(protocol_name.loc[0, 'footer_text'])} от {protocol_name.loc[0, 'date']}"
				else:
					errors.append(f"Для файла {doc_name} не нашли протокол, документ не заполнен")
				if len(conclusion_name) == 0:
					documents.loc[index_doc, 'text_conclusion'] = f"уч. № {str(add_conclusion_num)} от {add_conclusion_date}"
				else:
					documents.loc[index_doc, 'text_conclusion'] = f"уч. № {str(conclusion_name.loc[0, 'footer_text'])} от {conclusion_name.loc[0, 'date']}"
				# documents.loc[index_doc, 'executor'] = prescription
				executor = prescription
				documents.loc[index_doc, 'change_date'] = True
			elif re.findall(r'акт', doc_name.lower()):
				# documents.loc[index_doc, 'executor'] = act
				executor = act
				documents.loc[index_doc, 'change_date'] = True
			elif re.findall(r'утверждение', doc_name.lower()):
				# documents.loc[index_doc, 'executor'] = statement
				executor = statement
				documents.loc[index_doc, 'change_date'] = True
			else:
				logging.info('Документ не в списке необходимых, продолжаем')
				continue
			text_finish = f"Уч. № {footer_text}\nОтп. 1 экз. в адрес\n{hdd_number}\nИсп. {executor}\nПеч. {print}\n{date}\nб/ч"
			documents.loc[index_doc, 'date'] = date
			documents.loc[index_doc, 'first_header_text'] = text_first_header
			documents.loc[index_doc, 'footer_text'] = footer_text
			documents.loc[index_doc, 'text_finish'] = text_finish
			documents.loc[index_doc, 'executor'] = executor
			pages = pages_count(doc)
			page = 0
			if pages['error']:
				errors.append(f"Для файла {doc_name} подсчёт кол-ва страниц завершился с ошибкой")
			else:
				page = pages['text']
			documents.loc[index_doc, 'pages'] = page
			for dict_40_name in doc_name_for_dict_40:
				if dict_40_name in doc_name.lower():
					dict_40.append({doc_name: [classified, footer_text, num_scroll, str(page - 1)]})
			if not dict_file and all([True if _ not in doc_name.lower() else False for _ in doc_name_for_continue]):
				num_2 = str(int(num_2) + 1)  # Увеличиваем значение для учетного номера
			# percent_val += percent  # Увеличиваем прогресс
			# progress.emit(int(percent_val))  # Посылаем значение в прогресс бар
		if errors:
			return {'error': True, 'text': errors, 'data': ''}
		else:
			return {'error': False, 'text': '', 'data': {'num_1': num_1, 'num_2': num_2}}
	except BaseException as exception:
		return {'error': True, 'text': str(exception) + '\n' + traceback.format_exc(), 'data': {}}
