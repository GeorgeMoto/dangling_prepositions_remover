# Обработчик висячих предлогов

Утилита для автоматизированной обработки DOCX-документов с целью устранения "висячих предлогов" и улучшения типографики в русскоязычных текстах.

## Описание

Данное приложение заменяет обычные пробелы после предлогов и союзов на неразрывные, чтобы избежать ситуаций, когда предлог оказывается в конце строки, а слово, к которому он относится — в начале следующей строки. Также обрабатываются даты в формате "26 января 1994", где пробелы между числом, месяцем и годом заменяются на неразрывные.

### Основные возможности:

- Обработка одиночных DOCX-файлов
- Пакетная обработка всех DOCX-файлов в выбранной папке
- Настройка списка предлогов и союзов через пользовательский интерфейс
- Сохранение настроек в JSON-файле для последующего использования
- Подробное логирование процесса обработки
- Индикация прогресса обработки файлов

## Установка

### Требования

- Python 3.6+
- ttkbootstrap
- Стандартная библиотека Python (zipfile, re, json, threading, ...)

### Установка зависимостей

```bash
pip install ttkbootstrap
```

### Запуск программы

```bash
python main.py
```

## Использование

1. Запустите программу через `main.py`
2. Выберите один из вариантов:
   - **Выбрать файл** - для обработки одного DOCX-документа
   - **Выбрать папку** - для обработки всех DOCX-файлов в папке
3. Программа создаст папку `output_files` в той же директории, где находится исходный файл (или файлы)
4. Обработанные файлы будут сохранены в этой папке с оригинальными именами

### Настройка списка предлогов

1. Перейдите на вкладку "Предлоги"
2. Добавьте или удалите предлоги по вашему усмотрению
3. Нажмите "Сохранить изменения", чтобы сохранить список для последующего использования

## Структура проекта

- `main.py` - Главный модуль для запуска программы
- `ui.py` - Модуль пользовательского интерфейса
- `logic.py` - Модуль обработки DOCX-файлов
- `config.py` - Конфигурационный файл
- `prepositions.json` - Список предлогов и союзов (создается при первом запуске)
- `logs/` - Каталог с логами программы (создается автоматически)

## Как это работает

Программа использует следующий алгоритм:
1. Открывает DOCX-файл как ZIP-архив (DOCX - это ZIP-архив с XML-файлами)
2. Извлекает файл `word/document.xml`, содержащий текст документа
3. Выполняет поиск предлогов и дат в тексте с помощью регулярных выражений
4. Заменяет обычные пробелы на неразрывные (Unicode-символ 00A0)
5. Сохраняет измененный XML-файл обратно в DOCX

## Логирование

Программа ведет подробные логи работы, сохраняя их в папке `logs/`. Для каждого запуска создается отдельный лог-файл с датой и временем в названии.

## Устранение неполадок

### Проблема: Программа сообщает, что файл поврежден

Убедитесь, что файл является валидным DOCX-документом. Программа проверяет целостность файла перед обработкой.

### Проблема: Невозможно обработать файл из-за ошибки доступа

Убедитесь, что файл не открыт в других программах (например, в Microsoft Word).

## Лицензия

MIT

## Автор

Georgiy Motorin
