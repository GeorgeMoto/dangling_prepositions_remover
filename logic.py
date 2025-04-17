#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import zipfile
import re
import os
import tempfile
import shutil
from pathlib import Path
import logging

# Список предлогов, которые нужно обработать
PREPOSITIONS = {
    # Русские предлоги
    'в', 'во', 'на', 'к', 'ко', 'с', 'со', 'из', 'от', 'у',
    'о', 'об', 'про', 'за', 'над', 'под', 'при', 'без',
    'до', 'для', 'через', 'между', 'по', 'около', 'из-за', 'из-под',
    'не', 'то', 'и'
    # Английские предлоги были удалены, чтобы сосредоточиться на русских
}

# Список русских месяцев
MONTHS = {
    'января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
    'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря',
    'январь', 'февраль', 'март', 'апрель', 'май', 'июнь',
    'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'
}


def fix_hanging_prepositions_and_dates(input_file, output_file, prepositions=None, months=None, progress_callback=None):
    """
    Заменяет обычные пробелы после предлогов и в датах на неразрывные в DOCX документе.

    Args:
        input_file (str): Путь к исходному DOCX файлу
        output_file (str): Путь для сохранения обработанного файла
        prepositions (list, optional): Список предлогов для обработки. По умолчанию None (используется PREPOSITIONS).
        months (list, optional): Список месяцев для обработки дат. По умолчанию None (используется MONTHS).
        progress_callback (callable, optional): Функция обратного вызова для отображения прогресса.
                                              Принимает значение от 0.0 до 1.0.

    Returns:
        bool: True если обработка успешна, иначе False
    """
    # Проверяем, были ли предоставлены предлоги и месяцы, иначе используем значения по умолчанию
    if prepositions is None:
        prepositions = PREPOSITIONS

    if months is None:
        months = MONTHS

    # Проверяем входной файл
    input_path = Path(input_file)
    if not input_path.exists():
        logging.error(f"Ошибка: Файл не найден: {input_file}")
        raise FileNotFoundError(f"Файл не найден: {input_file}")

    if not input_path.suffix.lower() == '.docx':
        logging.error(f"Ошибка: Файл должен иметь расширение .docx: {input_file}")
        raise ValueError(f"Файл должен иметь расширение .docx: {input_file}")

    logging.info(f"Обработка файла: {input_file}")
    logging.info(f"Будет сохранено как: {output_file}")
    logging.info(f"Используемые предлоги: {', '.join(prepositions)}")
    logging.info(f"Обработка дат с месяцами: {', '.join(months)}")

    # Создаем временную директорию
    temp_dir = tempfile.mkdtemp()
    logging.info(f"Создана временная директория: {temp_dir}")

    try:
        # Проверяем, является ли файл валидным ZIP архивом (DOCX - это ZIP файл)
        try:
            with zipfile.ZipFile(input_file, 'r') as zip_check:
                # Проверяем, есть ли в нём основные компоненты DOCX
                file_list = zip_check.namelist()
                required_files = ['[Content_Types].xml', 'word/document.xml']
                for req_file in required_files:
                    if not any(f == req_file or f.endswith('/' + req_file) for f in file_list):
                        logging.error(
                            f"Ошибка: DOCX файл поврежден или имеет неверную структуру. Отсутствует {req_file}")
                        raise ValueError(f"DOCX файл поврежден или имеет неверную структуру. Отсутствует {req_file}")
        except zipfile.BadZipFile:
            logging.error(f"Ошибка: Файл {input_file} не является валидным DOCX файлом (поврежден ZIP архив)")
            raise ValueError(f"Файл {input_file} не является валидным DOCX файлом")

        # Сообщаем о прогрессе (10%)
        if progress_callback:
            progress_callback(0.1)

        # Распаковываем DOCX во временную директорию
        logging.info("Распаковка документа...")
        with zipfile.ZipFile(input_file, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # Сообщаем о прогрессе (20%)
        if progress_callback:
            progress_callback(0.2)

        # Путь к document.xml
        document_xml_path = os.path.join(temp_dir, 'word', 'document.xml')

        if not os.path.exists(document_xml_path):
            logging.error("Ошибка: Структура DOCX файла повреждена: не найден document.xml")
            raise ValueError("Структура DOCX файла повреждена: не найден document.xml")

        logging.info("Чтение и обработка document.xml...")
        # Читаем document.xml
        try:
            with open(document_xml_path, 'r', encoding='utf-8') as file:
                content = file.read()
        except UnicodeDecodeError:
            logging.error("Ошибка: Невозможно прочитать файл document.xml, возможно файл поврежден")
            raise ValueError("Невозможно прочитать файл document.xml, возможно файл поврежден")

        # Сообщаем о прогрессе (40%)
        if progress_callback:
            progress_callback(0.4)

        # Неразрывный пробел
        non_breaking_space = chr(160)  # Символ NO-BREAK SPACE (Unicode 00A0)

        # 1. Заменяем пробелы после предлогов на неразрывные
        pattern_prepositions = r'\b(' + '|'.join(map(re.escape, prepositions)) + r')\s'
        content, count_prepositions = re.subn(pattern_prepositions, r'\1' + non_breaking_space, content)
        logging.info(f"Заменено {count_prepositions} обычных пробелов после предлогов на неразрывные")

        # Сообщаем о прогрессе (60%)
        if progress_callback:
            progress_callback(0.6)

        # 2. Заменяем пробелы в датах формата "26 января 1994" на неразрывные
        # Создаем регулярное выражение для дат: число + пробел + месяц + пробел + год
        pattern_dates = r'(\b\d{1,2})\s(' + '|'.join(map(re.escape, months)) + r')\s(\d{4})\b'

        # В замене сохраняем число, месяц и год, но меняем обычные пробелы на неразрывные
        content, count_dates = re.subn(pattern_dates,
                                       r'\1' + non_breaking_space + r'\2' + non_breaking_space + r'\3',
                                       content)
        logging.info(f"Заменено {count_dates} обычных пробелов в датах на неразрывные")

        # Сообщаем о прогрессе (80%)
        if progress_callback:
            progress_callback(0.8)

        # Записываем обратно измененный document.xml
        with open(document_xml_path, 'w', encoding='utf-8') as file:
            file.write(content)

        logging.info("Создание нового DOCX файла...")
        # Создаем новый DOCX файл
        with zipfile.ZipFile(output_file, 'w') as outzip:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    outzip.write(file_path, arcname)

        logging.info(f"Документ успешно обработан и сохранен как {output_file}")
        logging.info(f"Всего заменено пробелов: {count_prepositions + count_dates * 2}")

        # Сообщаем о завершении (100%)
        if progress_callback:
            progress_callback(1.0)

        return True

    except Exception as e:
        logging.error(f"Ошибка при обработке файла: {str(e)}")
        # Выведем расширенную информацию об ошибке для отладки
        import traceback
        traceback.print_exc()
        raise e

    finally:
        # Удаляем временную директорию
        logging.info(f"Удаление временной директории: {temp_dir}")
        shutil.rmtree(temp_dir)