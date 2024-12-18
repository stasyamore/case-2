import os
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement

# Функция для изменения формата документа
def format_docx(file_path):
    doc = Document(file_path)
# Изменяем стиль каждого абзаца в документе
    for paragraph in doc.paragraphs:
        # Устанавливаем шрифт и размер шрифта
        for run in paragraph.runs:
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)
 # Устанавливаем межстрочный интервал на 1.5
        paragraph.paragraph_format.line_spacing = 1.5
   # cохраняем изменения
    doc.save(file_path)
    
# Путь к папке с документами
folder_path = 'path/to/your/documents'  
# Проходим по всем файлам в папке
for filename in os.listdir(folder_path):
    if filename.endswith('.docx'):
        file_path = os.path.join(folder_path, filename)
        format_docx(file_path)
        print(f"Форматирование завершено для: {filename}")
