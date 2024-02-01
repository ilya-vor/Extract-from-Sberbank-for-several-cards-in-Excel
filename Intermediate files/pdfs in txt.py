import fitz

def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        # Открываем PDF-файл для чтения бинарного содержимого
        doc = fitz.open(pdf_path)
        
        # Итерируем по страницам и записываем текст в текстовый файл
        for page_number in range(doc.page_count):
                page = doc[page_number]
                text += page.get_text()
        print(f'Текст успешно извлечен из PDF и сохранен в файл {pdf_path}')

    except Exception as e:
        print(f'Произошла ошибка: {str(e)}')

    return text

# Вызываем функцию для извлечения текста из PDF и записи его в текстовый файл
for i in range(1,5):
    try:
        pdf_file_path = f"{i}.pdf"
        text = extract_text_from_pdf(pdf_file_path)
        txt_path = f"{i}.txt"
        if text != "":
            with open(txt_path, 'w', encoding='utf-8') as txt_file:
                txt_file.write(text)
    except Exception as e:
        print(f'Произошла ошибка: {str(e)}')
