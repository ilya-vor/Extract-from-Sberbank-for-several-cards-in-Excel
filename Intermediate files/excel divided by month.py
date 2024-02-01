import openpyxl

# Открываем Excel файл
workbook = openpyxl.load_workbook('Выписка по дебетовой карте (на русском).xlsx')

Months = ["Декабрь","Ноябрь","Октябрь",
          "Сентябрь","Август","Июль",
          "Июнь","Май","Апрель",
          "Март","Февраль","Январь"]
Months = Months[::-1]

#Открываем общую историю операций по карте 
sheet = workbook['Sheet']

inews = [2 for i in range (12)]
iold = 2
while sheet[f"A{iold}"].value != None:
    cell_value = sheet[f"A{iold}"].value
    Number_of_mounth = int(str(cell_value)[3:5])
    Name_of_mounth = Months[Number_of_mounth - 1]
    try:
        target_sheet = workbook[Name_of_mounth]
    except:
        target_sheet = workbook.create_sheet(title=Name_of_mounth)
        for col in "ABCDE":
            target_sheet[f"{col}1"].value = sheet[f"{col}1"].value
        
    for col in "ABCDE":
        target_sheet[f"{col}{inews[Number_of_mounth - 1]}"].value = sheet[f"{col}{iold}"].value
    inews[Number_of_mounth - 1] += 1
    iold += 1

# Сохраняем файл
workbook.save('Выписка за год (по месяцам).xlsx')

print('Новый лист успешно создан и данные записаны.')
