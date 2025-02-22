from pydantic import BaseModel
from pathlib import Path


class Document(BaseModel):
	name: str  # Имя документа
	number: str  # Номер документа
	start_path: Path  # Используется как id, путь где лежит файл
	finish_path: Path  # Путь, где файл должен оказаться
	parent_path: Path  # Родительская директория
	secret_number: str  # Секретный номер
	action: str  # Действие над файлом
	conclusion: str  # Ссылка на заключение
	protocol: str  # Ссылка на протокол
	prescription: str  # Ссылка на предписание
	text: str  # Текст для вставки
	change_date: bool  # Вставка даты, если нужно
	executor: str  # Исполнитель документа
	first_header_text: str  # Текст для колонтитула первой страницы
	footer_text: str  # Текст для нижнего колонтитула
	date: str  # Дата для вставки
	text_conclusion: str  # Текст для ЗАКЛНОМ
	text_protocol: str  # Текст для ПРОТНОМ
	text_finish: str  # Текст для последней страницы
	pages: int  # Кол-во страниц в документе

