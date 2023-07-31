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
# Загружаем ваш файл в переменную `file` / вместо 'example' укажите название свого файла из текущей директории
#file = 'krasnodar — копия.xls'
# Загружаем spreadsheet в объект pandas
#xl = pd.ExcelFile(file)
# Печатаем название листов в данном файле
#print(xl.sheet_names)
# Загрузить лист в DataFrame по его имени: df1
#df1 = xl.parse('TDSheet')



# List all files and directories in current directory
#print(os.listdir('.'))

# Load in the workbook
#wb = load_workbook('./2.xlsx')

# Get sheet names
#print(wb.get_sheet_names())

# Get a sheet by name
#sheet = wb.get_sheet_by_name('TDSheet')

# Print the sheet title
#print(sheet.title)

# Get currently active sheet
#anotherSheet = wb.active

# Check `anotherSheet`
#print(anotherSheet)

#print(sheet['A1'].value)
# Select element 'B2' of your sheet
#c = sheet['B2']

# Retrieve the row number of your element
#print(c.value)
#print(c.row)

# Retrieve the column letter of your element
#print(c.column)

# Retrieve the coordinates of the cell
#print(c.coordinate)






#print('Начало работы')
#print('Магазин                  выручка')
#for i in range(7, 30):
    #if sheet.cell(row=i, column=1).value != None:
        #print(i, sheet.cell(row=i, column=1).value,'|', (sheet.cell(row=i, column=5).value))

#c = sheet.cell(row =sheet.max_row, column=sheet.max_column)
#print(c.coordinate)
# Вывести максимальное количество строк
#print(sheet.cell(row =sheet.max_row, column=sheet.max_column).value)

# Вывести максимальное количество колонок
#max_col = sheet.max_column
#max_Row = sheet.max_row

#df = pd.DataFrame(sheet.values)
#print(df)
#for r in dataframe_to_rows(df):
    #print(r)
# Put the sheet values in `data`
#data = sheet.values

# Indicate the columns in the sheet values
#cols = next(data)[1:]

# Convert your data to a list
#data = list(data)

# Read in the data at index 0 for the indices
#idx = [r[0] for r in data]

# Slice the data at index 1
#data = (islice(r, 1, None) for r in data)

# Make your DataFrame
#df1 = pd.DataFrame(data, index=idx, columns=cols)

#print('Принт df',df1)
#for r in dataframe_to_rows(df):
    #print(r)

# Import `xlwt`


# Initialize a workbook
#book = xlwt.Workbook(encoding="utf-8")

# Add a sheet to the workbook
#sheet1 = book.add_sheet("Python Sheet 1")

# Write to the sheet of the workbook
#sheet1.write(0, 0, "Начало анализа")



# Save the workbook
#book.save("spreadsheet.xls")