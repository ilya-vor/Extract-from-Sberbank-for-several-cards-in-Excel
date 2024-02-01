import fitz

def extract_text_from_pdf(pdf_path, txt_path):
    try:
        # Открываем PDF-файл для чтения бинарного содержимого
        doc = fitz.open(pdf_path)
        # Создаем объект для записи текста в текстовый файл
        with open(txt_path, 'w', encoding='utf-8') as txt_file:
            # Итерируем по страницам и записываем текст в текстовый файл
            for page_number in range(doc.page_count):
                page = doc[page_number]
                text = page.get_text()
                txt_file.write(text)

        print(f'Текст успешно извлечен из PDF и сохранен в файл: {txt_path}')

    except Exception as e:
        print(f'Произошла ошибка: {str(e)}')

# Задаем пути к файлам
pdf_file_path = '1.pdf'
txt_file_path = 'Выписка по дебетовой карте (на русском).txt'

# Вызываем функцию для извлечения текста из PDF и записи его в текстовый файл
extract_text_from_pdf(pdf_file_path, txt_file_path)


