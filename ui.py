import os
import ttkbootstrap as ttk
from tkinter import filedialog, messagebox
from logic import fix_hanging_prepositions  # Импортируем бизнес-логику
import logging


def log_separator():
    """Добавляет разделитель в лог-файл."""
    logging.info("---" * 30)


def select_file():
    """Выбор одного .docx файла."""
    try:
        file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if file_path:
            process_files([file_path])
    except Exception as e:
        logging.error(f"Ошибка при выборе файла: {e}", exc_info=True)
        messagebox.showerror("Ошибка", f"Не удалось выбрать файл: {e}")
        log_separator()


def select_folder():
    """Выбор папки с .docx файлами."""
    try:
        folder_path = filedialog.askdirectory()
        if folder_path:
            docx_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(".docx")]
            if not docx_files:
                messagebox.showwarning("Ошибка", "В папке нет .docx файлов")
                return
            process_files(docx_files)
    except Exception as e:
        logging.error(f"Ошибка при выборе папки: {e}", exc_info=True)
        messagebox.showerror("Ошибка", f"Не удалось выбрать папку: {e}")
        log_separator()


def process_files(files):
    """Обрабатывает файлы и сохраняет результат в output_files/."""
    if not files:
        return

    output_dir = os.path.join(os.path.dirname(files[0]), "output_files")
    os.makedirs(output_dir, exist_ok=True)

    errors = []  # Список для хранения ошибок

    for file_path in files:
        output_path = os.path.join(output_dir, os.path.basename(file_path))

        try:
            # Логируем начало обработки файла
            logging.info(f"Начало обработки файла: {file_path}")

            # Пытаемся обработать файл
            fix_hanging_prepositions(file_path, output_path)

            # Логируем успешное завершение
            logging.info(f"Файл успешно обработан: {output_path}")
            messagebox.showinfo("✅ Готово!", f"Файл сохранен в {output_path}")
        except Exception as e:
            error_message = f"Ошибка обработки файла {file_path}: {e}"
            errors.append(error_message)
            logging.error(error_message, exc_info=True)

        log_separator()

    # Если были ошибки, показываем их
    if errors:
        messagebox.showerror("Ошибка", "\n".join(errors))


# === Создание UI ===
def run_ui():
    root = ttk.Window(themename="superhero")  # Темная тема
    root.title("Обработка висячих предлогов")
    root.geometry("400x250")
    root.resizable(False, False)

    ttk.Label(root, text="Выберите файл или папку", font=("Arial", 14, "bold")).pack(pady=10)

    ttk.Button(root, text="📄 Выбрать файл", command=select_file, bootstyle="primary").pack(pady=5, padx=20, fill="x")
    ttk.Button(root, text="📁 Выбрать папку", command=select_folder, bootstyle="info").pack(pady=5, padx=20, fill="x")

    root.mainloop()