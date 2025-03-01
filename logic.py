import re
import os
from docx import Document
from contextlib import contextmanager
from config import SHORT_WORDS


@contextmanager
def safe_document_handling(input_path, output_path):
    """Контекстный менеджер для безопасной работы с документом."""
    # Проверка существования файла
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Файл {input_path} не найден.")

    # Проверка, открыт ли файл
    try:
        with open(input_path, 'rb') as _:
            pass
    except PermissionError:
        raise PermissionError(f"Файл {input_path} уже открыт в другой программе.")

    try:
        doc = Document(input_path)
        yield doc
        # Сохраняем файл только если все операции прошли успешно
        doc.save(output_path)
    except Exception as e:
        # Перебрасываем исключение наверх для обработки в UI
        raise e


def replace_spaces_with_nbsp(text):
    """Заменяет пробел после коротких предлогов и союзов на неразрывный пробел (\u00A0)."""
    words = re.split(r'(\s+)', text)  # Разбиваем по пробелам, но сохраняем их
    for i in range(len(words) - 2):
        word = words[i].strip(".,!?;:")
        if word.lower() in SHORT_WORDS and words[i + 1] == " ":
            words[i + 1] = "\u00A0"
    return "".join(words)


def process_paragraph(paragraph):
    """Обрабатывает текст в параграфе, заменяя пробелы после предлогов на неразрывные."""
    if not paragraph.runs:
        return
    for run in paragraph.runs:
        run.text = replace_spaces_with_nbsp(run.text)


def fix_hanging_prepositions(input_path, output_path, progress_callback=None):
    """
    Основная функция обработки документа.

    Args:
        input_path: Путь к исходному файлу
        output_path: Путь для сохранения обработанного файла
        progress_callback: Функция обратного вызова для обновления прогресса
    """
    with safe_document_handling(input_path, output_path) as doc:
        total_elements = (
                len(doc.paragraphs) +
                sum(len(row.cells) for table in doc.tables for row in table.rows) +
                sum(len(section.header.paragraphs) + len(section.footer.paragraphs) for section in doc.sections)
        )

        processed = 0

        # Обрабатываем обычные абзацы
        for paragraph in doc.paragraphs:
            process_paragraph(paragraph)
            processed += 1
            if progress_callback:
                progress_callback(processed / total_elements)

        # Обрабатываем текст в таблицах
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        process_paragraph(paragraph)
                        processed += 1
                        if progress_callback:
                            progress_callback(processed / total_elements)

        # Обрабатываем колонтитулы
        for section in doc.sections:
            for paragraph in section.header.paragraphs:
                process_paragraph(paragraph)
                processed += 1
                if progress_callback:
                    progress_callback(processed / total_elements)
            for paragraph in section.footer.paragraphs:
                process_paragraph(paragraph)
                processed += 1
                if progress_callback:
                    progress_callback(processed / total_elements)