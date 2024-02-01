import openpyxl

# Открываем Excel файл
workbook = openpyxl.load_workbook('Выписка за год (по месяцам, отсортированная).xlsx')

Months = ["Декабрь","Ноябрь","Октябрь",
          "Сентябрь","Август","Июль",
          "Июнь","Май","Апрель",
          "Март","Февраль","Январь"]
Months = Months[::-1]

for month in Months:
    print(month)
    sheet = workbook[month]
    sheet['J1'] = "Бюджет"
    sheet['J3'] = "Категории"
    sheet['K3'] = "Сумма"

    i = 2
    dohod = 0
    while any([sheet[f"G{ii}"].value for ii in range(i,i+3)]):
        if sheet[f"H{i}"].value != None:
            if sheet[f"H{i}"].value > 0:
                dohod += sheet[f"H{i}"].value
        i += 1

    sheet['J3'] = "Доход"
    sheet['K3'] = dohod

    i = 2
    while any([sheet[f"G{ii}"].value for ii in range(i,i+3)]):
        if sheet[f"G{i}"].value == "Прочие операции\n":
            ii = i + 1
            s = 0
            while sheet[f"G{ii}"].value != None:
                if sheet[f"H{ii}"].value < 0:
                    s += sheet[f"H{ii}"].value
                ii += 1
            sheet['J4'] = "Прочие операции"
            sheet['K4'] = s
            break
        i += 1
    
    i = 2
    ib = 5
    while any([sheet[f"G{ii}"].value for ii in range(i,i+3)]):
        if sheet[f"G{i}"].value == "Итого:":
            sheet[f'K{ib}'].value = sheet[f"H{i}"].value
            iname = i
            while sheet[f'G{iname}'].value != None:
                iname -= 1
            sheet[f'J{ib}'].value = sheet[f"G{iname+1}"].value
            ib += 1
        i += 1
    sheet[f"J{ib}"] = "Итого:"
    i = 3
    s = 0
    while sheet[f"K{i}"].value != None:
        s += float(sheet[f"K{i}"].value)
        i += 1
    sheet[f"K{i}"] = s
            
            
workbook.save('Выписка за год РЕЗУЛЬТАТ.xlsx')

print('Новый лист успешно создан и данные записаны.')
