#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import zipfile
import re
import os
import tempfile
import shutil
from pathlib import Path

# Настройте пути к файлам здесь
INPUT_FILE = "C:/Users/Георгий/OneDrive/Desktop/test.docx"  # Укажите путь к входному файлу
OUTPUT_FILE = "C:/Users/Георгий/OneDrive/Desktop/result.docx"  # Укажите путь к выходному файлу

# Список предлогов, которые нужно обработать
PREPOSITIONS = [
    # Русские предлоги
    'в', 'во', 'на', 'к', 'ко', 'с', 'со', 'из', 'от', 'у',
    'о', 'об', 'про', 'за', 'над', 'под', 'при', 'без',
    'до', 'для', 'через', 'между', 'по', 'около', 'из-за', 'из-под',
    'не', 'то', 'и'
    # Английские предлоги были удалены, чтобы сосредоточиться на русских
]


def fix_hanging_prepositions(input_file, output_file, prepositions):
    """
    Заменяет обычные пробелы после предлогов на неразрывные в DOCX документе.

    Args:
        input_file (str): Путь к исходному DOCX файлу
        output_file (str): Путь для сохранения обработанного файла
        prepositions (list): Список предлогов для обработки

    Returns:
        bool: True если обработка успешна, иначе False
    """
    # Проверяем входной файл
    input_path = Path(input_file)
    if not input_path.exists():
        print(f"Ошибка: Файл не найден: {input_file}")
        return False

    if not input_path.suffix.lower() == '.docx':
        print(f"Ошибка: Файл должен иметь расширение .docx: {input_file}")
        return False

    print(f"Обработка файла: {input_file}")
    print(f"Будет сохранено как: {output_file}")
    print(f"Используемые предлоги: {', '.join(prepositions)}")

    # Создаем временную директорию
    temp_dir = tempfile.mkdtemp()
    print(f"Создана временная директория: {temp_dir}")

    try:
        # Распаковываем DOCX во временную директорию
        print("Распаковка документа...")
        with zipfile.ZipFile(input_file, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # Путь к document.xml
        document_xml_path = os.path.join(temp_dir, 'word', 'document.xml')

        if not os.path.exists(document_xml_path):
            print("Ошибка: Структура DOCX файла повреждена: не найден document.xml")
            return False

        print("Чтение и обработка document.xml...")
        # Читаем document.xml
        with open(document_xml_path, 'r', encoding='utf-8') as file:
            content = file.read()

        # Создаем регулярное выражение для поиска предлогов со следующим пробелом
        # Используем \b для границы слова, чтобы находить только предлоги, а не части слов
        pattern = r'\b(' + '|'.join(map(re.escape, prepositions)) + r')\s'

        # Подсчитываем количество замен для отчета
        count_before = len(re.findall(pattern, content))

        # Заменяем пробелы после предлогов на неразрывные (Unicode символ NO-BREAK SPACE)
        # Используем прямую вставку символа, а не его экранирование через \u00A0
        non_breaking_space = chr(160)  # Символ NO-BREAK SPACE (Unicode 00A0)
        content = re.sub(pattern, r'\1' + non_breaking_space, content)

        # Подсчитываем оставшиеся совпадения для проверки
        remaining = len(re.findall(pattern, content))
        count_replaced = count_before - remaining

        print(f"Заменено {count_replaced} обычных пробелов на неразрывные")

        # Записываем обратно измененный document.xml
        with open(document_xml_path, 'w', encoding='utf-8') as file:
            file.write(content)

        print("Создание нового DOCX файла...")
        # Создаем новый DOCX файл
        with zipfile.ZipFile(output_file, 'w') as outzip:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    outzip.write(file_path, arcname)

        print(f"Документ успешно обработан и сохранен как {output_file}")
        return True

    except Exception as e:
        print(f"Ошибка при обработке файла: {str(e)}")
        # Выведем расширенную информацию об ошибке для отладки
        import traceback
        traceback.print_exc()
        return False

    finally:
        # Удаляем временную директорию
        print(f"Удаление временной директории: {temp_dir}")
        shutil.rmtree(temp_dir)


def main():
    """
    Основная функция для запуска обработки документа
    """
    success = fix_hanging_prepositions(INPUT_FILE, OUTPUT_FILE, PREPOSITIONS)

    if success:
        print("Обработка завершена успешно!")
    else:
        print("Обработка завершилась с ошибками.")


if __name__ == "__main__":
    main()