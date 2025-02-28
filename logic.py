import re
from docx import Document

# Список предлогов и союзов, которые не должны висеть в конце строки
SHORT_WORDS = {
    "в", "с", "на", "по", "за", "к", "у", "о", "об", "под", "над", "из", "от",
    "для", "про", "между", "через", "перед", "при", "до", "без", "вне", "кроме",
    "вместо", "со", "ко", "во", "и", "а", "но"
}

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

def fix_hanging_prepositions(input_path, output_path):
    """Основная функция обработки документа."""
    doc = Document(input_path)

    # Обрабатываем обычные абзацы
    for paragraph in doc.paragraphs:
        process_paragraph(paragraph)

    # Обрабатываем текст в таблицах
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    process_paragraph(paragraph)

    # Обрабатываем колонтитулы
    for section in doc.sections:
        for paragraph in section.header.paragraphs:
            process_paragraph(paragraph)
        for paragraph in section.footer.paragraphs:
            process_paragraph(paragraph)

    # Сохраняем файл
    doc.save(output_path)
    print(f"✅ Обработка завершена! Файл сохранен в {output_path}")
