# импорт библиотеки pandas
import pandas as pd

# Загружаем ваш файл в переменную `file` / вместо 'example' укажите название свого файла из текущей директории
file = 'example.xlsx'

# Загружаем spreadsheet в объект pandas
xl = pd.ExcelFile(file)

# Печатаем название листов в данном файле
print(xl.sheet_names)

# Загрузить лист в DataFrame по его имени: df1
df1 = xl.parse('Sheet1')

#writer = pd.ExcelWriter('krasnodar.xls', engine='xlsxwriter')

# Записать ваш DataFrame в файл
# a = yourData.to_excel(writer, 'TDSheet')

# Сохраним результат
#writer.save()