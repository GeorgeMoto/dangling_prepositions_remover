#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Главный модуль программы для запуска приложения обработки висячих предлогов в DOCX документах.
Настраивает логирование и запускает пользовательский интерфейс.
"""

import logging
import os
import sys
from datetime import datetime
from ui import Application
import ttkbootstrap as ttk


def setup_logging():
    """Настраивает логирование программы."""
    # Создаем директорию для логов, если ее нет
    logs_dir = "logs"
    if not os.path.exists(logs_dir):
        os.makedirs(logs_dir)

    # Создаем имя файла лога с датой и временем
    log_filename = os.path.join(logs_dir, f"app_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log")

    # Настраиваем логирование
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filename, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )

    logging.info("Запуск программы обработки висячих предлогов")
    logging.info(f"Логи сохраняются в: {log_filename}")


def main():
    """Основная функция запуска программы."""
    # Настраиваем логирование
    setup_logging()

    try:
        # Запускаем UI приложение
        logging.info("Инициализация пользовательского интерфейса")
        root = ttk.Window(themename="superhero")  # Темная тема
        root.title("Обработка висячих предлогов")
        app = Application(root)

        logging.info("Запуск главного цикла приложения")
        root.mainloop()

    except Exception as e:
        logging.critical(f"Критическая ошибка при запуске приложения: {e}", exc_info=True)
        print(f"Произошла ошибка: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()