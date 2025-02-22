import logging
import re
import shutil
import time
import traceback

import pythoncom
import fitz
import os
import pandas as pd
import numpy as np
from word2pdf import word2pdf
from zipfile import ZipFile
from pathlib import Path
from natsort import natsorted
from typing import Any


def sorting_files(path: Path, path_new: Path, fso: bool) -> dict:
	"""Сортирует файлы для подсчета значения прогресса и последующей вставки текстовки и номеров"""
	try:
		application_dict = {}  # Для подсчёта листов - новое.
		error_text = []
		data = {}
		if fso:
			docs = {}
			for folder in os.listdir(path):
				if 'проверке' in folder.lower():
					if folder != 'Материалы по специальной проверке технических средств':
						error_text.append('Название папки «Материалы по специальной проверке'
										  ' технических средств» написано с ошибками')
					else:
						file = os.listdir(Path(path, folder))
						file = natsorted(file, key=lambda y: y.rpartition(' ')[2][:-5])
						for name_element in ['акт', 'заключение']:
							for element in file:
								if name_element in element.lower():
									docs[element] = Path(path, folder)
				elif 'исследованиям' in folder.lower():
					if folder != 'Материалы по специальным исследованиям технических средств':
						error_text.append('Название папки «Материалы по специальным исследованиям'
										  ' технических средств» написано с ошибками')
					else:
						file = os.listdir(Path(path, folder))
						file = natsorted(file, key=lambda y: y.rpartition(' ')[2][:-5])
						for name_element in ['протокол', 'предписание']:
							for element in file:
								if name_element in element.lower():
									docs[element] = Path(path, folder)
				elif 'дополнительные' in folder.lower() and os.path.isdir(folder):
					if folder != 'Дополнительные материалы':
						error_text.append('Название папки «Дополнительные материалы» написано с ошибками')
					else:
						shutil.copytree(Path(path, folder), Path(path_new, folder))
			if error_text:
				return {'error': True, 'text': error_text, 'data': {}}
			for element in os.listdir():
				if 'сопроводит' in element.lower():
					docs[element] = path
			data['docx_for_progress'] = len(docs)
		else:
			docs = [file for file in os.listdir(path) if file.endswith('.docx')]  # Список документов
			docs = natsorted(docs, key=lambda y: y.rpartition(' ')[2][:-5])
			docs_ = [j_ for i_ in
					 ['^Акт', 'Приложение \d? к акту', '^Заключение', 'Приложение \d? к заключению', 'Протокол',
					  'Приложение А', 'Предписание', 'Форма 3', 'Опись', 'Сопроводит'] for j_ in docs if
					 re.findall(i_, j_, re.I)]
			docs_not = [i_ for i_ in docs if i_ not in docs_ and '~' not in i_]
			docs = docs_not + docs_
			# Процент для прогресса
			docx_for_progress = 0
			for name_file in os.listdir(path):
				if re.findall(r'приложение а', name_file.lower()):
					with ZipFile(Path(path, name_file)) as my_doc:
						xml_content = my_doc.read('docProps/app.xml')  # Общие свойства
						pages = int(re.findall(r'<Pages>(\w*)</Pages>', xml_content.decode())[0])
					if pages == 1:
						pythoncom.CoInitializeEx(0)
						# self.logging.info(f"Считаем кол-во листов в приложении {name_file}")
						# status.emit(f"Считаем кол-во листов в приложении {name_file}")
						word2pdf(str(Path(path, name_file)), str(Path(path, name_file + '.pdf')))
						input_file = fitz.open(str(Path(path, name_file + '.pdf')))  # Открываем
						pages = input_file.page_count  # Получаем кол-во страниц
						input_file.close()  # Закрываем
						os.remove(str(Path(path, name_file + '.pdf')))  # Удаляем pdf документ
					application_dict[name_file] = pages
				else:
					docx_for_progress += 1
			data['docs_for_progress'] = docx_for_progress
		data['docs'] = [Path(path, doc) for doc in docs]
		return {'error': False, 'text': '', 'data': data}
	except BaseException as exception:
		return {'error': True, 'text': str(exception) + '\n' + traceback.format_exc(), 'data': {}}


def report_rso(path: Path):
	"""Функция для генерации отчёта МВД"""
	file_mo = [mo for mo in os.listdir(path) if mo[-3:] == 'txt' and 'F19' in mo][0]
	df_report_rso = pd.read_csv(Path(path, file_mo), delimiter='|', encoding='ANSI', names=[
		'Порядковый номер лицензиата',
		'Серийный номер комплекта',
		'Серийный номер системного блока', 'удалить'])
	df_report_rso['Порядковый номер лицензиата'] = df_report_rso['Порядковый номер лицензиата'].astype(str)
	df_report_rso['№'] = np.arange(1, 1 + len(df_report_rso))
	df_report_rso = df_report_rso.reindex(columns=['№',
												   'Порядковый номер лицензиата',
												   'Серийный номер комплекта',
												   'Серийный номер системного блока',
												   'Заключение',
												   'Кол-во листов закл.',
												   'Протокол',
												   'Кол-во листов прот.',
												   'Предписание',
												   'Кол-во листов пред.',
												   'Сумма листов на комплект'])


def pages_count(file: Path) -> dict:
	"""Функция для подсчёта количества страниц в документе. Принимает путь к файлу"""
	# Конвертируем
	# while True:
	try:
		name = file.name
		parent_path = file.parent
		pythoncom.CoInitializeEx(0)
		name_pdf = name + '.pdf'
		word2pdf(str(Path(parent_path, name)), str(Path(parent_path, name_pdf)))
		input_file_pdf = fitz.open(str(Path(parent_path, name_pdf)))  # Открываем пдф
		count_page = input_file_pdf.page_count  # Получаем кол-во страниц
		input_file_pdf.close()  # Закрываем
		os.remove(str(Path(parent_path, name_pdf)))  # Удаляем пдф документ
		# temp_docx = os.path.join(parent_path, name)
		# temp_zip = os.path.join(parent_path, name + ".zip")
		# temp_folder = os.path.join(parent_path, "template")
		#
		# if os.path.exists(temp_zip):
		# 	rm(temp_zip)
		# if os.path.exists(temp_folder):
		# 	rm(temp_folder)
		# if os.path.exists(Path(parent_path, 'zip')):
		# 	rm(Path(parent_path, 'zip'))
		# os.rename(temp_docx, temp_zip)
		# os.mkdir(Path(parent_path, 'zip'))
		# with ZipFile(temp_zip) as my_document:
		# 	my_document.extractall(temp_folder)
		# pages_xml = os.path.join(temp_folder, "docProps", "app.xml")
		# string = open(pages_xml, 'r', encoding='utf-8').read()
		# string = re.sub(r"<Pages>(\w*)</Pages>",
		# 				"<Pages>" + str(count_page) + "</Pages>", string)
		# with open(pages_xml, "wb") as file_wb:
		# 	file_wb.write(string.encode("UTF-8"))
		# try_number = 0
		# while True:
		# 	try:
		# 		os.remove(temp_zip)
		# 		break
		# 	except PermissionError:
		# 		if try_number == 4:
		# 			break
		# 		time.sleep(3)
		# 		try_number += 1
		# if try_number == 4:
		# 	return {'error': True, 'text': 'Не удалось удалить файл'}
		# shutil.make_archive(temp_zip.replace(".zip", ""), 'zip', temp_folder)
		# os.rename(temp_zip, temp_docx)  # rename zip file to docx
		# rm(temp_folder)
		# rm(Path(parent_path, 'zip'))
		return {'error': False, 'text': count_page}
	except BaseException as exception:
		return {'error': True, 'text': f'Ошибка при подсчёте количества страниц: {exception}'}


def rm(folder_path):
	try:
		while len(os.listdir(folder_path)) != 0:
			time.sleep(0.5)
			for file_object in os.listdir(folder_path):
				flag_ = True
				while flag_:
					try:
						file_object_path = os.path.join(folder_path, file_object)
						try:
							if os.path.isfile(file_object_path) or os.path.islink(file_object_path):
								os.remove(file_object_path)
							else:
								try:
									shutil.rmtree(file_object_path)
								except FileNotFoundError:
									pass
						except OSError:
							os.remove(folder_path)
						flag_ = False
					except BaseException:
						pass
	except NotADirectoryError:
		os.remove(folder_path)
	time.sleep(0.05)
	shutil.rmtree(folder_path)


def return_error(log: logging, warning: str, status, status_text: str, default_path: Path, status_finish: list,
				 window, event: Any = False, error: str = '', info_value=None) -> None:
	"""Функция для возврата ошибок по всей программе и отмены оперции пользователем. """
	if info_value is None:
		info_value = []
	if event:
		log.error(error)
		if isinstance(info_value[2], list):
			error_text = '\n'.join(info_value[2])
		else:
			error_text = 'Работа программы завершена из-за непредвиденной ошибки, обратитесь к разработчику'
		info_value[0].emit(info_value[1], error_text)
		event.clear()
		event.wait()
		log.error(warning)
		status.emit(status_text)
		os.chdir(default_path)
		status_finish[0].emit(status_finish[1], status_finish[2])
		time.sleep(1)  # Не удалять, не успевает отработать emit status_finish. Может потом
		window.close()
		return
