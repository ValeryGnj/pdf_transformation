import PyPDF2
import os

def split_pdf_by_bookmarks(pdf_path, output_folder="output"):
    # Создаем папку для сохранения файлов, если ее еще нет
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Открываем PDF-файл
    with open(pdf_path, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)

        # Создаем список закладок и обращаем его
        bookmarks = pdf_reader.outline
        
        # Перебираем закладки
        for index, bookmark in enumerate(bookmarks):
            bookmark_title = bookmark.title
            print(bookmark_title)
            # Получаем номер страницы, связанной с закладкой
            start_page = pdf_reader.get_destination_page_number(bookmark)
            # print(start_page)

            # Определяем номер страницы следующей закладки
            if index < len(bookmarks) - 1:
                end_page = pdf_reader.get_destination_page_number(bookmarks[index + 1])
                # print(end_page)
            else:
                # Если текущая закладка последняя, берем все страницы до конца документа
                end_page = len(pdf_reader.pages)
                # print(end_page)

            # Создаем новый PDF-файл для каждой закладки
            output_pdf_writer = PyPDF2.PdfWriter()
            for page_number in range(start_page, end_page):
                output_pdf_writer.add_page(pdf_reader.pages[page_number - 1])                 
            # Сохраняем новый PDF-файл с именем, основанным на названии закладки
            output_file_path = os.path.join(output_folder, f"{bookmark_title}_document.pdf")
            with open(output_file_path, 'wb') as output_file:
                output_pdf_writer.write(output_file)
                print(f"Создан файл: {output_file_path}")
        
if __name__ == "__main__":
    pdf_path = "C:/Users/user/Desktop/pdf_transformation/certificates.pdf"
    output_folder = "C:/Users/user/Desktop/pdf_transformation/output_folder"

    split_pdf_by_bookmarks(pdf_path, output_folder)
