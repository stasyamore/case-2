## main стартовый модуль проекта

from functions import format_docx, change_format
import os

def main():
    # TODO вызов функции change_format()

    # TODO все что ниже переннести в change_format()

    # Путь к папке с документами
    folder_path = '/Users/dmitry.dobrozan/Desktop/case3'  
    # Проходим по всем файлам в папке
    for filename in os.listdir(folder_path):
        if filename.endswith('.docx'):
            file_path = os.path.join(folder_path, filename)
            format_docx(file_path)
            print(f"Форматирование завершено для: {filename}")

# инициализационный скрипт
if __name__ == "__main__":
    main()
    
