from docx import Document
from openpyxl import Workbook
import os

# Путь к папке с документами .docx (для примера используем текущую папку)
docx_folder_path = 'data\\'
excel_file_path = 'C:\\Users\\IMatveev\\PycharmProjects\\LKEDocxToXls\\venv\\documents.xlsx'
wb = Workbook()
ws = wb.active

# Ищем все .docx файлы в папке
docx_files = [f for f in os.listdir(docx_folder_path) if f.endswith('.docx')]

# Перебираем все .docx файлы и переносим их содержимое в Excel
for docx_file in docx_files:
    doc_path = os.path.join(docx_folder_path, docx_file)
    doc = Document(doc_path)

    # Итерация по объектам документа
    for element in doc.element.body:
        # Обработка параграфов
        if element.tag.endswith('p'):
            ws.append([p.text for p in element.xpath('.//w:t')])
        # Обработка таблиц
        elif element.tag.endswith('tbl'):
            for row in element.xpath('.//w:tr'):
                row_data = [cell.text for cell in row.xpath('.//w:tc/w:p/w:r/w:t')]
                ws.append(row_data)

# Сохраняем Excel файл
wb.save(excel_file_path)
# Возвращаем путь к созданному Excel файлу
#excel_file_path