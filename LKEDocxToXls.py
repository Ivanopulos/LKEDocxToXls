from docx import Document
import openpyxl
from openpyxl.utils import get_column_letter
import os

# Путь к папке с документами .docx (для примера используем текущую папку)
docx_folder_path = '/mnt/data/'
excel_file_path = '/mnt/data/documents.xlsx'

# Функция для извлечения текста и формул из параграфов документа Word
def extract_text_and_formulas(paragraphs):
    content = []
    for para in paragraphs:
        # Добавляем текст каждого параграфа
        content.append(para.text)
    return content

# Создаем новый Excel файл
wb = openpyxl.Workbook()
ws = wb.active

# Ищем все .docx файлы в заданной папке
docx_files = [f for f in os.listdir(docx_folder_path) if f.endswith('.docx')]

# Переменная для отслеживания номера строки в Excel файле
excel_row = 1

# Читаем каждый .docx файл и переносим содержимое в Excel
for docx_file in docx_files:
    doc_path = os.path.join(docx_folder_path, docx_file)
    doc = Document(doc_path)
    content = extract_text_and_formulas(doc.paragraphs)

    # Переносим содержимое в Excel, каждый параграф в новую ячейку в столбце A
    for para in content:
        col = 'A'  # Можно изменить, если нужно начать с другого столбца
        ws[f'{col}{excel_row}'] = para
        excel_row += 1

# Сохраняем Excel файл
wb.save(excel_file_path)

# Возвращаем путь к созданному Excel файлу
excel_file_path