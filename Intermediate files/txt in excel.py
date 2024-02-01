import openpyxl

# Открытие файла
with open('Выписка по дебетовой карте (на русском).txt', 'r', encoding='utf-8') as file:
    # Чтение содержимого файла построчно в список
    lines = file.readlines()

# Создаем новый Excel файл
workbook = openpyxl.Workbook()

# Выбираем активный лист
sheet = workbook.active

# Записываем текст в разные ячейки
sheet['A1'] = "Дата"
sheet['B1'] = "Время"
sheet['C1'] = "Категория"
sheet['D1'] = "Комментарий"
sheet['E1'] = "Сумма"
itxt = 30
iexcel = 2
while True:
    try:
        if lines[itxt] == "Продолжение на следующей странице\n":
            itxt += 11
        sheet[f"A{iexcel}"] = lines[itxt]
        sheet[f"B{iexcel}"] = lines[itxt+1]
        sheet[f"C{iexcel}"] = lines[itxt+4]
        sheet[f"D{iexcel}"] = lines[itxt+5][0:lines[itxt+5].index('.')]
        if lines[itxt+6][0] == "*":
            itxt += 1
        cleaned_string = ''.join(char for char in lines[itxt+6] if char.isdigit() or char in {'-', '.', '+', ','})
        cleaned_string = cleaned_string.replace(',','.')
        if cleaned_string[0] == '+':
            sheet[f"E{iexcel}"] = float(cleaned_string)
        else:
            sheet[f"E{iexcel}"] = float(cleaned_string) * -1
        itxt += 7
        iexcel += 1
    except Exception as e:
        print(f'Произошла ошибка: {str(e)}')
        break
for b in "ABCD":
    sheet[f"{b}{iexcel}"] = ""
# Сохраняем файл
workbook.save('Выписка по дебетовой карте (на русском).xlsx')

print('Файл успешно создан и сохранен.')
