# импорт библиотеки pandas
import pandas as pd
import aspose.slides as slides
import aspose.pydrawing as drawing
    # Загружаем ваш файл в переменную `file` / вместо 'example' укажите название свого файла из текущей директории
    #file = 'example.xlsx'

    # Загружаем spreadsheet в объект pandas
    #xl = pd.ExcelFile(file)

    # Печатаем название листов в данном файле
    #print(xl.sheet_names)

    # Загрузить лист в DataFrame по его имени: df1
    #df1 = xl.parse('Sheet1')

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
# Create presentation
with slides.Presentation('Анализ продаж по кварталам 1.pptx') as pres:
    # Access first slide
    slide = pres.slides[0]

    # Access first slide
    sld = pres.slides[0]

    # Add chart with default data
    chart = sld.shapes.add_chart(slides.charts.ChartType.CLUSTERED_COLUMN, 0, 0, 500, 500)

    # Set chart Title
    chart.chart_title.add_text_frame_for_overriding("Sample Title")
    #chart.chart_title.text_frame_for_overriding.text_frame_format.center_text = 1
    chart.chart_title.height = 20
    chart.has_title = True

    # Set first series to Show Values
    chart.chart_data.series[0].labels.default_data_label_format.show_value = True

    # Set the index of chart data sheet
    defaultWorksheetIndex = 0

    # Get the chart data worksheet
    fact = chart.chart_data.chart_data_workbook

    # Delete default generated series and categories
    chart.chart_data.series.clear()
    chart.chart_data.categories.clear()
    s = len(chart.chart_data.series)
    s = len(chart.chart_data.categories)

    # Add new series
    chart.chart_data.series.add(fact.get_cell(defaultWorksheetIndex, 0, 1, "Series 1"), chart.type)
    chart.chart_data.series.add(fact.get_cell(defaultWorksheetIndex, 0, 2, "Series 2"), chart.type)

    # Add new categories
    chart.chart_data.categories.add(fact.get_cell(defaultWorksheetIndex, 1, 0, "Caetegoty 1"))
    chart.chart_data.categories.add(fact.get_cell(defaultWorksheetIndex, 2, 0, "Caetegoty 2"))
    chart.chart_data.categories.add(fact.get_cell(defaultWorksheetIndex, 3, 0, "Caetegoty 3"))

    # Take first chart series
    series = chart.chart_data.series[0]

    # Now populating series data

    series.data_points.add_data_point_for_bar_series(fact.get_cell(defaultWorksheetIndex, 1, 1, 20))
    series.data_points.add_data_point_for_bar_series(fact.get_cell(defaultWorksheetIndex, 2, 1, 50))
    series.data_points.add_data_point_for_bar_series(fact.get_cell(defaultWorksheetIndex, 3, 1, 30))

    # Set fill color for series
    series.format.fill.fill_type = slides.FillType.SOLID
    series.format.fill.solid_fill_color.color = drawing.Color.red

    # Take second chart series
    series = chart.chart_data.series[1]

    # Now populating series data
    series.data_points.add_data_point_for_bar_series(fact.get_cell(defaultWorksheetIndex, 1, 2, 30))
    series.data_points.add_data_point_for_bar_series(fact.get_cell(defaultWorksheetIndex, 2, 2, 10))
    series.data_points.add_data_point_for_bar_series(fact.get_cell(defaultWorksheetIndex, 3, 2, 60))

    # Setting fill color for series
    series.format.fill.fill_type = slides.FillType.SOLID
    series.format.fill.solid_fill_color.color = drawing.Color.orange

    # First label will be show Category name
    lbl = series.data_points[0].label
    lbl.data_label_format.show_category_name = True

    lbl = series.data_points[1].label
    lbl.data_label_format.show_series_name = True

    # Show value for third label
    lbl = series.data_points[2].label
    lbl.data_label_format.show_value = True
    lbl.data_label_format.show_series_name = True
    lbl.data_label_format.separator = "/"

    # Save the presentation
    pres.save("create-chart-in-presentation.pptx", slides.export.SaveFormat.PPTX)