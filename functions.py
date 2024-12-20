# functions модуль для импорта функций для решения задачи  
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
            font = run.font
            font.name = 'Times New Roman'
            font.size = Pt(14)
        
        # Устанавливаем межстрочный интервал на 1.5
        paragraph.paragraph_format.line_spacing_rule = 1.5
    
    # Сохраняем изменения
    doc.save(file_path)


def change_format(folder_path):
    # Проходим по всем файлам в папке
    for filename in os.listdir(folder_path):
        if filename.endswith('.docx'):
            file_path = os.path.join(folder_path, filename)
            format_docx(file_path)
            print(f"Форматирование завершено для: {filename}")
