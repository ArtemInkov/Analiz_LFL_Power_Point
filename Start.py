"""Импорт os для адаптации путей
Импорт pandas для работы с файлами эксель
xlsxwriter для записи xl файлов
openpyxl import load_workbook для чтение xl файлов
import xlrd Библиотека для чтения и форматирования данных xls или xlsx
import xlwt
from openpyxl.utils.dataframe import dataframe_to_rows"""
import os
from itertools import islice
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
import xlsxwriter
from openpyxl import load_workbook
import xlrd
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.workbook import Workbook
import xlwt
# Загружаем ваш файл в переменную `file` / вместо 'example' укажите название свого файла из текущей директории
#file = 'krasnodar — копия.xls'
# Загружаем spreadsheet в объект pandas
#xl = pd.ExcelFile(file)
# Печатаем название листов в данном файле
#print(xl.sheet_names)
# Загрузить лист в DataFrame по его имени: df1
#df1 = xl.parse('TDSheet')



# List all files and directories in current directory
print(os.listdir('.'))

# Load in the workbook
wb = load_workbook('./2.xlsx')

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
print('Начало работы')

analize = xlwt.Workbook(encoding="utf-8") #Создаем переменную с книгой
sheet_analize_LFL = analize.add_sheet("Анализ год к году") #Создаем лист
data_sale = load_workbook('./2.xlsx') # Присваеваем файл с данными продаж к переменной data_sale


sheet = data_sale.get_sheet_by_name('TDSheet')
#sheet_data_sale = data_sale.wb.get_sheet_by_name('TDSheet') #Присваеваем лист  с данными продаж к переменной sheet_data_sale из файла data_sale
max_col = sheet.max_column  #записываем в переменную максимальное число колонок
max_Row = sheet.max_row #записываем в переменную максимальное число строк
last_coordinate = sheet.cell(row =sheet.max_row, column=sheet.max_column) # присваеваем координаты последней ячейки переменной
for cellObj in sheet['A1':last_coordinate.coordinate]:
    for cell in cellObj:
        #print(cell.coordinate, cell.value)
        if cell.value == 'Магазин':
            print(cell.column)
            print(cell.row)
            sheet_analize_LFL.write(0, 0, cell.value)
            for i in range(1, max_col):
                print(i, sheet.cell(row=cell.row, column=i).value)
                #print(cell.row)
                #print(sheet.cell(row=7, column=4).value)
                #sheet_analize_LFL.write(0, cell.column, mounth)
analize.save("анализ продаж.xls")

#for i in range(7, 30):
    #if sheet.cell(row=i, column=1).value != None:
        #print(i, sheet.cell(row=i, column=1).value,'|', (sheet.cell(row=i, column=5).value))
