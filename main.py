import os
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
# Изменяем стиль каждого абзаца в документе
for paragraph in doc.paragraphs:
        # Устанавливаем шрифт и размер шрифта
    for run in paragraph.runs:
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
# Путь к папке с документами
folder_path = 'path/to/your/documents'  
# Проходим по всем файлам в папке
for filename in os.listdir(folder_path):
    if filename.endswith('.docx'):
        file_path = os.path.join(folder_path, filename)
        format_docx(file_path)
        print(f"Форматирование завершено для: {filename}")