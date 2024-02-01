import openpyxl

# Открываем Excel файл
workbook = openpyxl.load_workbook('Выписка по дебетовой карте (на русском).xlsx')

# Выбираем активный лист
sheet = workbook.active

# Редактируем транзакции
iexcel = 2
while sheet[f"A{iexcel}"].value != None:
    if sheet[f"D{iexcel}"].value in ["SBOL",
                                     "SBERBANK ONL@IN VKLAD-KARTA",
                                     "ZACHISLENIE KREDITA",
                                     "SBERBANK ONL@IN KARTA-VKLAD",
                                     "BRANCH KARTA-KREDIT"]:
        sheet.delete_rows(iexcel)
    if str(sheet[f"D{iexcel}"].value)[0:10] == "Автоплатёж":
        sheet[f"C{iexcel}"] = "Комунальные платежи, связь, интернет."
    iexcel += 1

# Сохраняем файл
workbook.save('Выписка по дебетовой карте (на русском).xlsx')

print('Файл успешно создан и сохранен.')
