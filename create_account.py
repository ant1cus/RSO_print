import pandas as pd


def create_account(documents: pd.DataFrame, inventory_count: str, num_1: str, num_2: str):
	"""Создание описей и добавление их к текущим документам"""

	def sort(len_, d):  # Ф-я для записи в необходимом порядке
		if len_ <= 40:  # Если одной описи хватает для записи документов
			d.append(len_)  # Добавляем длину последних
			return d  # Возвращаем
		else:  # Если не хватает
			d.append(40)  # Добавляем 40 штук
			sort(len_ - 40, d)  # Рекурсия
	inventory = 1  # Если выбрана опись
	if inventory_count == 40:
		buff = []
		for el in dict_40:
			for i_ in el:
				if re.findall('Заключение', i_) or re.findall('Предписание', i_):
					buff.append(el)
		dict_40 = buff
		len_dict = int(len(dict_40) / 2)  # Получаем длину для записи в опись
		dict_for_op = []  # Список
		sort(len_dict, dict_for_op)  # Устанавливаем порядок
		dict_after = []  # Для записи
		start_ = 0  # Счетчик
		for el in dict_for_op:  # Для элементов
			if flag_inventory == 40:  # Если по 40 в одной описи
				for i_ in range(0, el):  # Заключения
					dict_after.append(dict_40[start_ + i_])
				for i_ in range(0, el):  # Предписания
					dict_after.append(dict_40[start_ + len_dict + i_])
			start_ += el
	else:
		dict_after = dict_40  # Если все в одной описи
	flag_for_op = 0
	percent = 10 / len(dict_after)
	logging.info(dict_after)
	for el in dict_after:  # Для получившихся элементов
		value = el.popitem()  # Забираем элемент
		status.emit('Добавляем документ ' + str(value[0]) + ' в опись')
		name_count = '\\Опись №' + str(inventory) + '.docx'
		if flag_for_op == 0:  # Если элемент первый в данной описи
			text_for_foot = str(num_1) + str(num_2) + 'c'
			document = docx.Document()  # Открываем
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
			section.different_first_page_header_footer = True
			# Добавляем необходимые надписи перед таблицей, выравниваем, создаем таблицу
			p = document.add_paragraph('Опись документов № ' + str(inventory))
			p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
			table = document.add_table(rows=1, cols=5, style='Table Grid')
			style = document.styles['Normal']
			font = style.font
			font.name = 'TimesNewRoman'
			font.size = Pt(12)
			cell_write(document.styles['Normal'], ['Порядковый номер', 'Наименование документа',
												   'Регистрационный номер',
												   'Номер экземпляра, гриф секретности',
												   'Количество листов в экземпляре'])
			# Текст внизу таблицы
			p = document.add_paragraph()
			p.text = '\n\n' + account_post + '\t\t\t\t\t\t\t\t' + account_signature
			p.paragraph_format.widow_control = True  # Чтобы подпись не убегала одна
			p.paragraph_format.keep_together = True  # Чтобы подпись не убегала одна
			logging.info("Вставляем колонтитул")
			insert_header(document, 11, value[1][0] + '\n(без приложения не секретно)\nЭкз.№ 1',
						  text_for_foot, hdd_number, executor,
						  print_people, date, account_path, name_count, fso)
			flag_for_op = 1  # Чтобы не создавать, если это не нужно
		# Открываем необходимую опись
		document = docx.Document(os.path.abspath(account_path + name_count))
		table = document.tables[0]  # Выбираем таблицу
		table.add_row()  # Добавляем колонку и значения
		style = document.styles['Normal']
		font = style.font
		font.name = 'TimesNewRoman'
		font.size = Pt(12)
		cell_write(document.styles['Normal'], [str(flag_for_op), value[0][:-5], value[1][1],
											   '№' + value[1][2] + ', ' + value[1][0], value[1][3]],
				   flag_for_op)
		flag_for_op += 1  # Увеличиваем счетчик

		document.save(account_path + name_count)  # Сохраняем документ
		if flag_inventory == 40:  # Если в описи по 40 штук
			if flag_for_op == 81:  # Если добавили все документы
				flag_for_op = 0  # Для создания новой описи
				num_2 = str(int(num_2) + 1)  # Увеличиваем значение для учетного номера
				inventory += 1  # Номер описи
		percent_val += percent  # Увеличиваем прогресс
		flag += 1  # Для того, что бы не мелькал прогресс бар
		if flag == 4:  # Только для каждого 4 документа при добавлении
			self.progress.emit(int(percent_val))  # Обновляем прогресс бар
			flag = 0