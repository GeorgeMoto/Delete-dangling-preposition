# ui.py
import os
import threading
import ttkbootstrap as ttk
from tkinter import filedialog, messagebox, StringVar
from ttkbootstrap.constants import *
from config import SHORT_WORDS, save_short_words
from logic import fix_hanging_prepositions
from logger import log_separator  # Импортируем функцию разделителя
import logging


class Application:
    def __init__(self, root):
        self.root = root
        self.root.title("Обработка висячих предлогов")
        self.root.geometry("550x600")
        self.root.resizable(True, True)

        # Переменные для отслеживания прогресса
        self.progress_var = ttk.DoubleVar()
        self.status_var = StringVar(value="Готов к работе")
        self.files_processed = 0
        self.total_files = 0

        # Создаем интерфейс
        self.create_ui()

    def create_ui(self):
        """Создает элементы пользовательского интерфейса."""
        # Создаем основной контейнер с вкладками
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=BOTH, expand=YES, padx=10, pady=10)

        # Вкладка обработки файлов
        files_frame = ttk.Frame(notebook)
        notebook.add(files_frame, text="Обработка файлов")

        # Вкладка настроек
        settings_frame = ttk.Frame(notebook)
        notebook.add(settings_frame, text="Настройки")

        # Наполняем вкладку обработки файлов
        self.build_files_tab(files_frame)

        # Наполняем вкладку настроек
        self.build_settings_tab(settings_frame)

        # Статусбар внизу приложения
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill=X, pady=(0, 10), padx=10)

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

    def build_files_tab(self, parent):
        """Создает интерфейс вкладки обработки файлов."""
        # Заголовок
        ttk.Label(
            parent,
            text="Выберите файл или папку для обработки",
            font=("Arial", 14, "bold")
        ).pack(pady=10)

        # Фрейм для кнопок
        buttons_frame = ttk.Frame(parent)
        buttons_frame.pack(fill=X, padx=20, pady=10)

        # Кнопки выбора файлов
        ttk.Button(
            buttons_frame,
            text="📄 Выбрать файл",
            command=self.select_file,
            bootstyle="primary",
            width=20
        ).pack(side=LEFT, padx=5)

        ttk.Button(
            buttons_frame,
            text="📁 Выбрать папку",
            command=self.select_folder,
            bootstyle="info",
            width=20
        ).pack(side=LEFT, padx=5)

        # Информационная панель
        info_frame = ttk.LabelFrame(parent, text="Информация")
        info_frame.pack(fill=BOTH, expand=YES, padx=20, pady=10)

        ttk.Label(
            info_frame,
            text=(
                "Программа заменяет обычные пробелы после предлогов и союзов\n"
                "на неразрывные пробелы, чтобы избежать 'висячих предлогов'\n"
                "в конце строки. Обработанные файлы сохраняются в папке 'output_files'."
            ),
            font=("Arial", 10),
            justify="left"
        ).pack(pady=10, padx=10)

    def build_settings_tab(self, parent):
        """Создает интерфейс вкладки настроек."""
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
        list_frame = ttk.LabelFrame(main_frame, text="Список предлогов и союзов")
        list_frame.pack(side=LEFT, fill=BOTH, expand=YES, padx=(0, 10))

        # Создаем список для отображения предлогов
        self.words_listbox = ttk.Treeview(
            list_frame,
            columns=("word",),
            show="headings",
            height=15
        )
        self.words_listbox.heading("word", text="Слово")
        self.words_listbox.column("word", width=100)
        self.words_listbox.pack(fill=BOTH, expand=YES, padx=5, pady=5)

        # Скроллбар для списка
        scrollbar = ttk.Scrollbar(
            list_frame,
            orient=VERTICAL,
            command=self.words_listbox.yview
        )
        scrollbar.pack(side=RIGHT, fill=Y)
        self.words_listbox.configure(yscrollcommand=scrollbar.set)

        # Загружаем данные в список
        for word in sorted(SHORT_WORDS):
            self.words_listbox.insert("", END, values=(word,))

        # Правая часть - кнопки управления
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(side=LEFT, fill=Y, padx=(0, 5))

        # Поле ввода для нового слова
        self.new_word_var = StringVar()
        ttk.Label(
            buttons_frame,
            text="Новое слово:"
        ).pack(anchor=W, pady=(0, 5))

        ttk.Entry(
            buttons_frame,
            textvariable=self.new_word_var,
            width=15
        ).pack(fill=X, pady=(0, 10))

        # Кнопки управления списком
        ttk.Button(
            buttons_frame,
            text="Добавить",
            command=self.add_word,
            bootstyle="success",
            width=15
        ).pack(fill=X, pady=5)

        ttk.Button(
            buttons_frame,
            text="Удалить выбранное",
            command=self.delete_word,
            bootstyle="danger",
            width=15
        ).pack(fill=X, pady=5)

        ttk.Separator(buttons_frame, orient=HORIZONTAL).pack(fill=X, pady=10)

        ttk.Button(
            buttons_frame,
            text="Сохранить изменения",
            command=self.save_changes,
            bootstyle="primary",
            width=15
        ).pack(fill=X, pady=5)

    def add_word(self):
        """Добавляет новое слово в список."""
        word = self.new_word_var.get().strip().lower()
        if not word:
            messagebox.showwarning("Предупреждение", "Введите слово для добавления")
            return

        # Проверяем, что слово еще не в списке
        existing_words = [self.words_listbox.item(item)["values"][0] for item in self.words_listbox.get_children()]
        if word in existing_words:
            messagebox.showwarning("Предупреждение", f"Слово '{word}' уже есть в списке")
            return

        # Добавляем в визуальный список и в набор
        self.words_listbox.insert("", END, values=(word,))
        SHORT_WORDS.add(word)
        self.new_word_var.set("")  # Очищаем поле ввода

    def delete_word(self):
        """Удаляет выбранное слово из списка."""
        selected = self.words_listbox.selection()
        if not selected:
            messagebox.showwarning("Предупреждение", "Выберите слово для удаления")
            return

        word = self.words_listbox.item(selected[0])["values"][0]
        self.words_listbox.delete(selected[0])
        SHORT_WORDS.discard(word)

    def save_changes(self):
        """Сохраняет изменения в списке слов."""
        if save_short_words(SHORT_WORDS):
            messagebox.showinfo("Успех", "Изменения сохранены")
        else:
            messagebox.showerror("Ошибка", "Не удалось сохранить изменения")

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
                docx_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(".docx")]
                if not docx_files:
                    messagebox.showwarning("Предупреждение", "В папке нет .docx файлов")
                    return
                self.process_files(docx_files)
        except Exception as e:
            logging.error(f"Ошибка при выборе папки: {e}", exc_info=True)
            messagebox.showerror("Ошибка", f"Не удалось выбрать папку: {e}")

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

    def process_worker(self, files, output_dir):
        """Рабочая функция для обработки файлов в отдельном потоке."""
        errors = []
        successful_files = 0

        for i, file_path in enumerate(files):
            output_path = os.path.join(output_dir, os.path.basename(file_path))

            try:
                # Логируем начало обработки файла
                logging.info(f"Начало обработки файла: {file_path}")

                # Обновляем UI перед началом обработки файла
                self.root.after(0, lambda i=i: self.status_var.set(f"Обработка файла {i + 1} из {len(files)}"))

                # Пытаемся обработать файл с обратным вызовом для прогресса
                fix_hanging_prepositions(
                    file_path,
                    output_path,
                    lambda p, i=i: self.root.after(0, lambda: self.update_progress(i, p, len(files)))
                )

                # Увеличиваем счетчик обработанных файлов
                successful_files += 1

                # Логируем успешное завершение
                logging.info(f"Файл успешно обработан: {output_path}")

            except FileNotFoundError as e:
                error_message = f"Файл не найден: {file_path}"
                errors.append(error_message)
                logging.error(error_message)

            except PermissionError as e:
                error_message = f"Нет доступа к файлу или файл открыт: {file_path}"
                errors.append(error_message)
                logging.error(error_message)

            except Exception as e:
                error_message = f"Ошибка обработки файла {file_path}: {str(e)}"
                errors.append(error_message)
                logging.error(error_message, exc_info=True)

        # Обновляем UI после завершения всех файлов
        self.root.after(0, lambda: self.processing_complete(successful_files, errors))

    def update_progress(self, file_index, file_progress, total_files):
        """Обновляет прогресс обработки файлов."""
        # Вычисляем общий прогресс: (текущий индекс файла + прогресс текущего) / всего файлов
        overall_progress = (file_index + file_progress) / total_files
        self.progress_var.set(overall_progress * 100)

        # Обновляем статус
        current_file = file_index + 1
        self.status_var.set(f"Обработка файла {current_file} из {total_files} ({int(file_progress * 100)}%)")

        # Обновляем UI
        self.root.update_idletasks()

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
            else:
                # Если все файлы с ошибками
                messagebox.showerror("Ошибка", f"Не удалось обработать файлы. Ошибок: {len(errors)}")
        else:
            # Если всё успешно
            messagebox.showinfo("✅ Готово!", f"Успешно обработано файлов: {successful_files}")

        # Сбрасываем статус и прогресс
        self.status_var.set("Готов к работе")
        self.progress_var.set(0)

    def process_files(self, files):
        """Обрабатывает файлы и сохраняет результат в output_files/."""
        if not files:
            return

        # Настраиваем каталог для выходных файлов
        output_dir = os.path.join(os.path.dirname(files[0]), "output_files")
        os.makedirs(output_dir, exist_ok=True)

        # Инициализируем переменные прогресса
        self.progress_var.set(0)
        self.files_processed = 0
        self.total_files = len(files)

        # Обновляем статус
        self.status_var.set(f"Подготовка к обработке {len(files)} файлов...")

        # Запускаем обработку в отдельном потоке
        worker_thread = threading.Thread(
            target=self.process_worker,
            args=(files, output_dir),
            daemon=True
        )
        worker_thread.start()


def run_ui():
    """Запускает пользовательский интерфейс."""
    root = ttk.Window(themename="superhero")  # Темная тема
    app = Application(root)
    root.mainloop()