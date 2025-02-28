import os
import ttkbootstrap as ttk
from tkinter import filedialog, messagebox
from logic import fix_hanging_prepositions  # Импортируем бизнес-логику


def select_file():
    """Выбор одного .docx файла."""
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if file_path:
        process_files([file_path])


def select_folder():
    """Выбор папки с .docx файлами."""
    folder_path = filedialog.askdirectory()
    if folder_path:
        docx_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(".docx")]
        if not docx_files:
            messagebox.showwarning("Ошибка", "В папке нет .docx файлов")
            return
        process_files(docx_files)


def process_files(files):
    """Обрабатывает файлы и сохраняет результат в output_files/."""
    if not files:
        return

    output_dir = os.path.join(os.path.dirname(files[0]), "output_files")
    os.makedirs(output_dir, exist_ok=True)

    for file_path in files:
        output_path = os.path.join(output_dir, os.path.basename(file_path))
        fix_hanging_prepositions(file_path, output_path)

    messagebox.showinfo("✅ Готово!", f"Файлы сохранены в {output_dir}")


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
