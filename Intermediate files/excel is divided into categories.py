import openpyxl

# Открываем Excel файл
workbook = openpyxl.load_workbook('Выписка за год (по месяцам).xlsx')

Months = ["Декабрь","Ноябрь","Октябрь",
          "Сентябрь","Август","Июль",
          "Июнь","Май","Апрель",
          "Март","Февраль","Январь"]
Months = Months[::-1]

for month in Months:
    sheet = workbook[month]
    sheet['G1'] = "Категории"
    categories = []
    i = 2
    icat = 2
    while sheet[f"C{i}"].value != None:
        if sheet[f"C{i}"].value not in categories:
            categories.append(sheet[f"C{i}"].value)
            commentaries = []
            while any([sheet[f"G{ii}"].value for ii in range(icat,icat+5)]):
                icat += 1
            icat += 1
            sheet[f"G{icat}"].value  = sheet[f"C{i}"].value 
            iccat = 2
            iname = icat
            while sheet[f"C{iccat}"].value != None:
                if sheet[f"C{iccat}"].value == categories[-1]:
                    if sheet[f"D{iccat}"].value in commentaries:
                        icccat = iname

                        while sheet[f"D{iccat}"].value != sheet[f"G{icccat}"].value:
                            icccat += 1

                            
                        sheet[f"H{icccat}"].value = float(sheet[f"H{icccat}"].value) + float(sheet[f"E{iccat}"].value)
                    else:
                        icat += 1
                        for col in "DE":
                            target_col = chr(ord(col) + 3)
                            sheet[f"{target_col}{icat}"].value = sheet[f"{col}{iccat}"].value
                        commentaries.append(sheet[f"D{iccat}"].value)
                iccat += 1
            if sheet[f"G{iname}"].value not in ["Прочие операции\n","Перевод с карты\n","Перевод на карту\n"]:
                iii = iname + 1
                s = 0
                while sheet[f"G{iii}"].value != None:
                    iii += 1
                    s += float(sheet[f"H{iii-1}"].value)
                sheet[f"G{iii}"] = "Итого:"
                sheet[f"H{iii}"] = s
        i += 1
            
workbook.save('Выписка за год (по месяцам, отсортированная).xlsx')

print('Новый лист успешно создан и данные записаны.')
