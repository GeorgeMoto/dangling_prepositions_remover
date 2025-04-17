import os
import threading
import zipfile
import ttkbootstrap as ttk
from tkinter import filedialog, messagebox, StringVar
from ttkbootstrap.constants import *
import logging
from pathlib import Path
import json

# Импортируем функцию из logic.py
from logic import fix_hanging_prepositions_and_dates, MONTHS

# Файл для хранения списка предлогов
PREPOSITIONS_FILE = "prepositions.json"


# Загрузка списка предлогов из JSON-файла
def load_prepositions():
    try:
        if os.path.exists(PREPOSITIONS_FILE):
            with open(PREPOSITIONS_FILE, 'r', encoding='utf-8') as f:
                return set(json.load(f))
        else:
            # Значения по умолчанию, если файл не существует
            default_prepositions = {
                'в', 'во', 'на', 'к', 'ко', 'с', 'со', 'из', 'от', 'у',
                'о', 'об', 'про', 'за', 'над', 'под', 'при', 'без',
                'до', 'для', 'через', 'между', 'по', 'около', 'из-за', 'из-под',
                'не', 'то', 'и'
            }
            # Сохраняем значения по умолчанию в файл
            save_prepositions(default_prepositions)
            return default_prepositions
    except Exception as e:
        logging.error(f"Ошибка при загрузке предлогов: {e}", exc_info=True)
        return set()


# Сохранение списка предлогов в JSON-файл
def save_prepositions(prepositions_set):
    try:
        with open(PREPOSITIONS_FILE, 'w', encoding='utf-8') as f:
            json.dump(list(prepositions_set), f, ensure_ascii=False, indent=4)
        return True
    except Exception as e:
        logging.error(f"Ошибка при сохранении предлогов: {e}", exc_info=True)
        return False


class Application:
    def __init__(self, root):
        self.root = root
        self.root.title("Обработка висячих предлогов и дат")
        self.root.geometry("600x650")  # Увеличил размер окна
        self.root.resizable(True, True)

        # Переменные для отслеживания прогресса
        self.progress_var = ttk.DoubleVar()
        self.status_var = StringVar(value="Готов к работе")
        self.files_processed = 0
        self.total_files = 0

        # Загружаем список предлогов из JSON
        self.prepositions = load_prepositions()
        self.months = list(MONTHS)

        # Создаем интерфейс
        self.create_ui()

    def create_ui(self):
        """Создает элементы пользовательского интерфейса."""
        # Создаем статусбар внизу приложения
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill=X, side=BOTTOM, pady=10, padx=10)

        # Прогресс-бар
        self.progress_bar = ttk.Progressbar(
            status_frame,
            variable=self.progress_var,
            mode="determinate",
            bootstyle="success"
        )
        self.progress_bar.pack(fill=X, side=TOP, pady=(0, 5))

        # Статус текст
        status_label = ttk.Label(
            status_frame,
            textvariable=self.status_var,
            font=("Arial", 10)
        )
        status_label.pack(side=LEFT, padx=5)

        # Создаем основной контейнер с вкладками
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=BOTH, expand=YES, padx=10, pady=(10, 0))

        # Вкладка обработки файлов
        files_frame = ttk.Frame(notebook)
        notebook.add(files_frame, text="Обработка файлов")

        # Вкладка настроек предлогов
        prepositions_frame = ttk.Frame(notebook)
        notebook.add(prepositions_frame, text="Предлоги")

        # Наполняем вкладку обработки файлов
        self.build_files_tab(files_frame)

        # Наполняем вкладку настроек предлогов
        self.build_prepositions_tab(prepositions_frame)

    def build_files_tab(self, parent):
        """Создает интерфейс вкладки обработки файлов."""
        # Фрейм для кнопок выбора файлов
        file_selection_frame = ttk.Frame(parent)
        file_selection_frame.pack(fill=X, padx=20, pady=10)

        # Кнопка выбора одного файла
        file_button = ttk.Button(
            file_selection_frame,
            text="Выбрать файл",
            command=self.select_file,
            bootstyle="primary",
            width=20
        )
        file_button.pack(side=LEFT, padx=5)

        # Кнопка выбора папки
        folder_button = ttk.Button(
            file_selection_frame,
            text="Выбрать папку",
            command=self.select_folder,
            bootstyle="secondary",
            width=20
        )
        folder_button.pack(side=LEFT, padx=5)

        # Информационный текст
        info_frame = ttk.LabelFrame(parent, text="Информация")
        info_frame.pack(fill=BOTH, expand=YES, padx=20, pady=10)

        ttk.Label(
            info_frame,
            text=(
                "Программа заменяет обычные пробелы после предлогов и союзов\n"
                "на неразрывные пробелы, чтобы избежать 'висячих предлогов'\n"
                "в конце строки. Также обрабатываются даты, заменяя пробелы в\n"
                "сочетаниях числа и месяца на неразрывные. Обработанные файлы\n"
                "сохраняются в папке 'output_files'."
            ),
            font=("Arial", 10),
            justify="left"
        ).pack(pady=10, padx=10)

        # Добавляем дополнительную информацию о текущем процессе
        self.process_info_frame = ttk.LabelFrame(parent, text="Информация о процессе")
        self.process_info_frame.pack(fill=X, padx=20, pady=(0, 10))

        self.process_info = ttk.Label(
            self.process_info_frame,
            text="Готов к обработке файлов",
            font=("Arial", 10),
            justify="left"
        )
        self.process_info.pack(pady=10, padx=10, fill=X)

    def build_prepositions_tab(self, parent):
        """Создает интерфейс вкладки настроек предлогов."""
        # Заголовок
        ttk.Label(
            parent,
            text="Настройка списка предлогов и союзов",
            font=("Arial", 14, "bold")
        ).pack(pady=10)

        # Фрейм для списка и кнопок
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=BOTH, expand=YES, padx=20, pady=10)

        # Левая часть - список слов
        list_frame = ttk.LabelFrame(main_frame, text="Предлоги и союзы")
        list_frame.pack(side=LEFT, fill=BOTH, expand=YES, padx=(0, 10))

        # Создаем список для отображения слов
        self.prepositions_listbox = ttk.Treeview(
            list_frame,
            columns=("word",),
            show="headings",
            height=15
        )
        self.prepositions_listbox.heading("word", text="Слово")
        self.prepositions_listbox.column("word", width=100)
        self.prepositions_listbox.pack(fill=BOTH, expand=YES, padx=5, pady=5)

        # Скроллбар для списка
        scrollbar = ttk.Scrollbar(
            list_frame,
            orient=VERTICAL,
            command=self.prepositions_listbox.yview
        )
        scrollbar.pack(side=RIGHT, fill=Y)
        self.prepositions_listbox.configure(yscrollcommand=scrollbar.set)

        # Загружаем данные в список
        for word in sorted(self.prepositions):
            self.prepositions_listbox.insert("", END, values=(word,))

        # Правая часть - кнопки управления
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(side=LEFT, fill=Y, padx=(0, 5))

        # Поле ввода для нового слова
        self.new_preposition_var = StringVar()
        ttk.Label(
            buttons_frame,
            text="Новое слово:"
        ).pack(anchor=W, pady=(0, 5))

        ttk.Entry(
            buttons_frame,
            textvariable=self.new_preposition_var,
            width=15
        ).pack(fill=X, pady=(0, 10))

        # Кнопки управления списком
        ttk.Button(
            buttons_frame,
            text="Добавить",
            command=self.add_preposition,
            bootstyle="success",
            width=15
        ).pack(fill=X, pady=5)

        ttk.Button(
            buttons_frame,
            text="Удалить выбранное",
            command=self.delete_preposition,
            bootstyle="danger",
            width=15
        ).pack(fill=X, pady=5)

        ttk.Separator(buttons_frame, orient=HORIZONTAL).pack(fill=X, pady=10)

        ttk.Button(
            buttons_frame,
            text="Сохранить изменения",
            command=self.save_prepositions,
            bootstyle="primary",
            width=15
        ).pack(fill=X, pady=5)

    # Методы для работы с предлогами
    def add_preposition(self):
        """Добавляет новый предлог в список."""
        word = self.new_preposition_var.get().strip().lower()
        if not word:
            messagebox.showwarning("Предупреждение", "Введите слово для добавления")
            return

        # Проверяем, что слово еще не в списке
        if word in self.prepositions:
            messagebox.showwarning("Предупреждение", f"Слово '{word}' уже есть в списке")
            return

        # Добавляем в визуальный список и в набор
        self.prepositions_listbox.insert("", END, values=(word,))
        self.prepositions.add(word)
        self.new_preposition_var.set("")  # Очищаем поле ввода

    def delete_preposition(self):
        """Удаляет выбранный предлог из списка."""
        selected = self.prepositions_listbox.selection()
        if not selected:
            messagebox.showwarning("Предупреждение", "Выберите слово для удаления")
            return

        word = self.prepositions_listbox.item(selected[0])["values"][0]
        self.prepositions_listbox.delete(selected[0])
        self.prepositions.remove(word)

    def save_prepositions(self):
        """Сохраняет изменения в списке предлогов в JSON файл."""
        if save_prepositions(self.prepositions):
            messagebox.showinfo("Успех", "Список предлогов сохранен в файл")
        else:
            messagebox.showerror("Ошибка", "Не удалось сохранить список предлогов")

    def select_file(self):
        """Выбор одного .docx файла."""
        try:
            file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
            if file_path:
                self.process_files([file_path])
        except Exception as e:
            logging.error(f"Ошибка при выборе файла: {e}", exc_info=True)
            messagebox.showerror("Ошибка", f"Не удалось выбрать файл: {e}")

    def select_folder(self):
        """Выбор папки с .docx файлами."""
        try:
            folder_path = filedialog.askdirectory()
            if folder_path:
                logging.info(f"Выбрана папка: {folder_path}")

                # Обновляем информацию о процессе
                self.process_info.config(text=f"Поиск .docx файлов в папке...\nПожалуйста, подождите")
                self.root.update_idletasks()

                # Получаем список .docx файлов
                docx_files = self.get_docx_files(folder_path)

                if not docx_files:
                    messagebox.showwarning("Предупреждение", "В папке нет .docx файлов")
                    self.process_info.config(text=f"Файлы .docx не найдены")
                    return

                # Показываем количество найденных файлов
                self.process_info.config(text=f"Найдено файлов .docx: {len(docx_files)}\nНачинаем обработку...")
                self.root.update_idletasks()

                self.process_files(docx_files)
        except Exception as e:
            logging.error(f"Ошибка при выборе папки: {e}", exc_info=True)
            messagebox.showerror("Ошибка", f"Не удалось выбрать папку: {e}")
            self.process_info.config(text=f"Ошибка: {str(e)}")

    def get_docx_files(self, folder_path):
        """Получает список .docx файлов из папки (без рекурсии)."""
        docx_files = []

        # Только файлы в выбранной папке (без подпапок)
        docx_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path)
                      if os.path.isfile(os.path.join(folder_path, f)) and f.endswith(".docx")]

        # Логируем общее количество найденных файлов
        logging.info(f"Всего найдено файлов: {len(docx_files)}")

        return docx_files

    def update_progress(self, file_progress):
        """Обновляет прогресс обработки файлов."""
        # Вычисляем общий прогресс: (завершенные файлы + прогресс текущего) / всего файлов
        overall_progress = (self.files_processed + file_progress) / self.total_files
        self.progress_var.set(overall_progress * 100)

        # Обновляем статус
        current_file = self.files_processed + 1
        self.status_var.set(f"Обработка файла {current_file} из {self.total_files} ({int(file_progress * 100)}%)")

        # Обновляем UI
        self.root.update_idletasks()

    def process_worker(self, files):
        """Рабочая функция для обработки файлов в отдельном потоке."""
        errors = []
        successful_files = 0

        # Получаем общую директорию для всех файлов
        folder_path = os.path.dirname(files[0]) if len(files) == 1 else os.path.dirname(os.path.commonpath(files))

        # Создаем основную директорию для выходных файлов
        output_dir = os.path.join(folder_path, "output_files")
        os.makedirs(output_dir, exist_ok=True)

        logging.info(f"Папка с исходными файлами: {folder_path}")
        logging.info(f"Папка для выходных файлов: {output_dir}")

        for i, file_path in enumerate(files):
            try:
                logging.info(f"Начало обработки файла: {file_path}")

                # Обновляем информацию о текущем файле
                current_file_info = f"Обработка файла {i + 1} из {len(files)}:\n{os.path.basename(file_path)}"
                self.root.after(0, lambda info=current_file_info: self.process_info.config(text=info))
                self.root.after(0, lambda i=i: self.status_var.set(f"Обработка файла {i + 1} из {len(files)}"))

                # Обновляем прогресс перед началом обработки файла
                current_progress = i / len(files)
                self.root.after(0, lambda p=current_progress: self.progress_var.set(p * 100))

                # Определяем выходной путь
                output_path = os.path.join(output_dir, os.path.basename(file_path))
                logging.info(f"Выходной путь: {output_path}")

                # Проверяем, можно ли открыть файл как DOCX
                try:
                    # Это дополнительная проверка на возможность открытия файла
                    with zipfile.ZipFile(file_path, 'r') as zip_check:
                        # Проверяем наличие необходимых файлов в DOCX
                        if not 'word/document.xml' in zip_check.namelist():
                            raise ValueError("Файл DOCX поврежден или имеет неверную структуру")
                except (zipfile.BadZipFile, zipfile.LargeZipFile):
                    raise ValueError(
                        f"Файл {os.path.basename(file_path)} не является валидным DOCX файлом или поврежден")

                # Вызываем функцию обработки с колбэком прогресса
                def progress_callback(file_progress):
                    # Пересчитываем общий прогресс: часть от текущего файла
                    file_part = 1.0 / len(files)
                    overall_progress = (i / len(files)) + (file_progress * file_part)
                    self.root.after(0, lambda p=overall_progress:
                    self.status_var.set(f"Обработка файла {i + 1} из {len(files)} ({int(file_progress * 100)}%)"))
                    self.root.after(0, lambda p=overall_progress: self.progress_var.set(p * 100))

                # Вызываем функцию обработки
                fix_hanging_prepositions_and_dates(
                    file_path,
                    output_path,
                    self.prepositions,
                    self.months,
                    progress_callback
                )

                self.files_processed += 1
                successful_files += 1
                logging.info(f"Файл успешно обработан: {output_path}")

                # Обновляем информацию о успешной обработке
                success_info = f"Файл успешно обработан:\n{os.path.basename(file_path)}"
                self.root.after(0, lambda info=success_info: self.process_info.config(text=info))

            except FileNotFoundError:
                error_message = f"Файл не найден: {file_path}"
                errors.append(error_message)
                logging.error(error_message)
                # Обновляем информацию об ошибке
                self.root.after(0, lambda msg=error_message: self.process_info.config(text=f"Ошибка: {msg}"))

            except PermissionError:
                error_message = f"Нет доступа к файлу или файл открыт: {file_path}"
                errors.append(error_message)
                logging.error(error_message)
                # Обновляем информацию об ошибке
                self.root.after(0, lambda msg=error_message: self.process_info.config(text=f"Ошибка: {msg}"))

            except ValueError as e:
                error_message = f"Ошибка формата файла: {str(e)}"
                errors.append(error_message)
                logging.error(error_message)
                # Обновляем информацию об ошибке
                self.root.after(0, lambda msg=error_message: self.process_info.config(text=f"Ошибка: {msg}"))

            except Exception as e:
                error_message = f"Ошибка обработки файла {os.path.basename(file_path)}: {str(e)}"
                errors.append(error_message)
                logging.error(error_message, exc_info=True)
                # Обновляем информацию об ошибке
                self.root.after(0, lambda msg=error_message: self.process_info.config(text=f"Ошибка: {msg}"))

        self.root.after(0, lambda: self.processing_complete(successful_files, errors))

    def processing_complete(self, successful_files, errors):
        """Вызывается после завершения обработки всех файлов."""
        # Показываем сообщение о результатах обработки
        if len(errors) > 0:
            # Если есть ошибки, но были успешные файлы
            if successful_files > 0:
                message = f"Обработано успешно: {successful_files} файлов\n\nОшибки ({len(errors)}):\n"
                # Показываем первые 3 ошибки, чтобы не перегружать окно
                for i, error in enumerate(errors[:3]):
                    message += f"\n{i + 1}. {error}"
                if len(errors) > 3:
                    message += f"\n\n...и ещё {len(errors) - 3} ошибок. Подробности в лог-файле."

                messagebox.showwarning("Обработка завершена с ошибками", message)
                # Обновляем информацию о процессе
                self.process_info.config(
                    text=f"Обработка завершена: {successful_files} успешно, {len(errors)} с ошибками")
            else:
                # Если все файлы с ошибками
                messagebox.showerror("Ошибка", f"Не удалось обработать файлы. Ошибок: {len(errors)}")
                # Обновляем информацию о процессе
                self.process_info.config(text=f"Обработка завершена с ошибками: {len(errors)} файлов")
        else:
            # Если всё успешно
            messagebox.showinfo("✅ Готово!", f"Успешно обработано файлов: {successful_files}")
            # Обновляем информацию о процессе
            self.process_info.config(text=f"Обработка успешно завершена: {successful_files} файлов")

        # Сбрасываем статус и прогресс
        self.status_var.set("Готов к работе")
        self.progress_var.set(0)
        self.files_processed = 0
        self.total_files = 0

    def process_files(self, files):
        """Обрабатывает файлы и сохраняет результат в output_files/."""
        if not files:
            return

        # Инициализируем переменные прогресса
        self.progress_var.set(0)
        self.files_processed = 0
        self.total_files = len(files)

        # Обновляем статус и информацию о процессе
        self.status_var.set(f"Подготовка к обработке {len(files)} файлов...")
        self.process_info.config(text=f"Подготовка к обработке {len(files)} файлов...")

        # Запускаем обработку в отдельном потоке
        worker_thread = threading.Thread(
            target=self.process_worker,
            args=(files,),
            daemon=True
        )
        worker_thread.start()


def run_ui():
    """Запускает пользовательский интерфейс."""
    # Настраиваем логирование
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler("app.log"),
            logging.StreamHandler()
        ]
    )

    root = ttk.Window(themename="superhero")  # Темная тема
    app = Application(root)
    root.mainloop()


if __name__ == "__main__":
    run_ui()