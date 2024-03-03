import os
import re
from PyPDF2 import PdfMerger
import pandas as pd
from glob import glob

def extract_last_name(pdf_filename):
    # Извлекаем фамилию из имени файла
    # Предполагаем, что фамилия идет первой в названии файла
    match = re.match(r'^([^\d]+)', pdf_filename)
    if match:
        last_name = match.group(1)
        # Удаляем недопустимые символы для имени файла
        cleaned_last_name = re.sub(r'[^\w\s]', '', last_name)
        return cleaned_last_name
    else:
        # Если не удается извлечь фамилию, просто возвращаем имя файла
        return os.path.splitext(pdf_filename)[0]

def merge_pdfs(pdf_list, output_filename):
    merger = PdfMerger()
    for pdf in pdf_list:
        try:
            merger.append(pdf)
        except FileNotFoundError:
            print(f"Warning: File not found - {pdf}")
    merger.write(output_filename)
    merger.close()

def read_excel(file_path, sheet_name, column_name, pdf_folder):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    file_dict = {row[column_name]: os.path.join(pdf_folder, row['Фамилия']) for _, row in df.iterrows()}
    return df, file_dict

excel_file_path = 'SR.xlsx'
excel_sheet_name = 'Лист1'
excel_column_name = 'Фамилия'
pdf_folder = "C:/Users/user/Desktop/pdf_transformation/output_folder"
output_folder = "C:/Users/user/Desktop/pdf_transformation/"

# Создаем папку для сохранения объединенного PDF-файла, если её нет
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

df, pdf_file_dict = read_excel(excel_file_path, excel_sheet_name, excel_column_name, pdf_folder)

output_pdf_path = os.path.join(output_folder, 'объединенный_файл.pdf')

# Получаем отсортированный список фамилий из Excel
sorted_last_names = sorted(pdf_file_dict.keys(), key=lambda x: df[df[excel_column_name] == x].index[0])

pdf_files_in_order = []
for last_name in sorted_last_names:
    # Поиск соответствующего PDF-файла
    matching_pdfs = glob(os.path.join(pdf_folder, f'*{last_name}*.pdf'))
    if matching_pdfs:
        pdf_files_in_order.append(matching_pdfs[0])

merge_pdfs(pdf_files_in_order, output_pdf_path)

print(f'PDF-файлы успешно объединены в файл {output_pdf_path}')
