from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import ChartData
from pptx.chart.data import CategoryChartData
from pptx.util import Inches
from pptx.dml.color import RGBColor
import pptx.enum.chart
from pptx.enum.chart import XL_TICK_MARK
from pptx.util import Pt
from openpyxl import load_workbook
import aspose.slides as slides
import aspose.pydrawing as drawing
from pptx.enum.chart import XL_LABEL_POSITION
from pptx.enum.chart import XL_LEGEND_POSITION
from pptx.enum.text import  MSO_AUTO_SIZE
import xlrd
a = 'анализ продаж.xlsx'
b = 'Анализ продаж по кварталам.pptx'
def presentation_pptx_ferst(file_name = ''):
    """Функция создает файл продажи и первый слайд берет года из файла exel"""
    file_xls = xlrd.open_workbook(file_name)
    sheets_xls = file_xls.sheet_by_name("Анализ год к году")
    last_year = sheets_xls.cell(1, 1).value.split()
    last_year = last_year[1]
    year = sheets_xls.cell(1, 2).value.split()
    year = year[1]
    presentation = Presentation()
    #help(presentation)
    print(presentation.slides)
    ferst_slide = presentation.slide_layouts[0]
    #help(ferst_slide)
    slide = presentation.slides.add_slide(ferst_slide)
    slide.shapes.title.text = f'Продажи к сравнению\n' \
                              f'{last_year} - {year}'

    presentation.save('Анализ продаж по кварталам.pptx')



# p = Presentation('Продажи 1 квартал  Краснодар.pptx')
# slide = p.slides.add_slide(p.slide_layouts[1])
# for shape in slide.placeholders:
#     print('%d %s' % (shape.placeholder_format.idx, shape.name))
# for shape in slide.shapes:
#     print('%s' % shape.shape_type)
# for shape in slide.shapes:
#     if shape.is_placeholder:
#         phf = shape.placeholder_format
#         print('%d, %s' % (phf.idx, phf.type))
#         if not shape.has_text_frame:
#             continue
#         text_frame = shape.text_frame

def presentation_pptx_shops_graphic(file_name_xls = '', file_name_pptx = ''):
    file_xls = xlrd.open_workbook(file_name_xls)
    sheets_xls = file_xls.sheet_by_name("Анализ год к году")
    colum_sh = sheets_xls.ncols
    row_sh = sheets_xls.nrows
    last_year = sheets_xls.cell(1, 1).value.split()
    last_year = last_year[1]
    year = sheets_xls.cell(1, 2).value.split()
    year = year[1]
    p = Presentation(file_name_pptx)

    print(range(int(colum_sh - 1)//4))
    for i in range(1, int((colum_sh - 1)//4)+ 1):
        print(int((colum_sh - 1)//4))
        print(i)
        if i == 1:
            mount_title = sheets_xls.cell(1, 1).value.split()
        else:
            mount_title = sheets_xls.cell(1, (i * 4)-3).value.split()
        print(mount_title)
        p.slide_width = Inches(15)
        ferst_slide = p.slide_layouts[5]

        # help(ferst_slide)
        slide = p.slides.add_slide(ferst_slide)
        s = p.slide_width
        print(s)
        slide.shapes.title.text = f'Продажи к сравнению в рублях\n{mount_title[0]}'
        #slide.placeholders[1].text = f'Продажи к сравнению\n' \
        #                              f'{mount_title}'
        # define chart data ---------------------
        chart_data = ChartData()
        shop_list = []
        last_year_value = []
        current_year_value = []
        procent_year = []
        summ_last_year_rub = []
        summ_current_year_rub = []
        summ_last_year_weight = []
        summ_current_year_weight = []

        for i_shop in range(1, row_sh):
            value_cell = sheets_xls.cell(i_shop, 0).value
            if value_cell != '' and value_cell != 'Магазин':
                shop_list.append(sheets_xls.cell(i_shop, 0).value)
                print(sheets_xls.cell(i_shop, 0).value)
                if i == 1:
                    #print(sheets_xls.cell(i_shop, 1).value)
                    #print(sheets_xls.cell(i_shop, 2).value)
                    if isinstance(sheets_xls.cell(i_shop, 1).value, float):
                        last_year_value.append(sheets_xls.cell(i_shop, 1).value)
                    else:
                        last_year_value.append(0)
                    if isinstance(sheets_xls.cell(i_shop, 2).value, float):
                        current_year_value.append(sheets_xls.cell(i_shop, 2).value)
                    else:
                        current_year_value.append(0)
                    if isinstance(sheets_xls.cell(i_shop, 1).value, float) and isinstance(sheets_xls.cell(i_shop, 2).value, float):
                        procent_year.append(round((100 - (sheets_xls.cell(i_shop, 1).value / sheets_xls.cell(i_shop, 2).value * 100)), 2))
                    else:
                        procent_year.append(0)
                else:
                    a1 = sheets_xls.cell(i_shop, (i * 4)-3).value
                    a = sheets_xls.cell(i_shop, (i * 4)-2).value
                    if isinstance(a1, float):
                        last_year_value.append(sheets_xls.cell(i_shop, (i * 4)-3).value)
                    else:
                        last_year_value.append(0)
                    if isinstance(a, float):
                        current_year_value.append(sheets_xls.cell(i_shop, (i * 4)-2).value)
                    else:
                        current_year_value.append(0)
                    if isinstance(a1, float) and isinstance(a, float):
                        procent_year.append(round(100 - (sheets_xls.cell(i_shop, (i * 4)-3).value / sheets_xls.cell(i_shop, (i * 4)-2).value * 100), 2))
                    else:
                        procent_year.append(0)
            print(shop_list)
            print(last_year_value)
            print(current_year_value)
            print(procent_year)

        chart_data.categories = shop_list
        chart_data.add_series(last_year, (tuple(last_year_value)))
        chart_data.add_series(year, (tuple(current_year_value)))
        chart_data.add_series('%', (tuple(procent_year)))

        summ_last_year_rub.append(sum(last_year_value))
        summ_current_year_rub.append(sum(current_year_value))



        x, y, cx, cy = Inches(0.01), Inches(2), Inches(15), Inches(4.5)
        graphic_frame = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data)
        chart = graphic_frame.chart
        plot = chart.plots[0]
        plot.has_data_labels = True
        data_labels = plot.data_labels

        data_labels.font.size = Pt(13)
        data_labels.font.color.rgb = RGBColor(0x0A, 0x42, 0x80)
        data_labels.position = XL_LABEL_POSITION.INSIDE_END
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.RIGHT
        chart.legend.include_in_layout = False

        """
        Создание слайда с продажами вес
        
        """

        p.slide_width = Inches(15)
        ferst_slide = p.slide_layouts[5]

        # help(ferst_slide)
        slide = p.slides.add_slide(ferst_slide)
        s = p.slide_width
        print(s)
        slide.shapes.title.text = f'Продажи к сравнению вес\n{mount_title[0]}'
        # slide.placeholders[1].text = f'Продажи к сравнению\n' \
        #                              f'{mount_title}'
        # define chart data ---------------------
        chart_data = ChartData()
        shop_list = []
        last_year_value = []
        current_year_value = []
        procent_year = []
        for i_shop in range(3, row_sh):
            value_cell = sheets_xls.cell(i_shop, 0).value
            if value_cell != '' and value_cell != 'Магазин':
                shop_list.append(sheets_xls.cell(i_shop, 0).value)
                print(sheets_xls.cell(i_shop, 0).value)
                if i == 3:
                    print(sheets_xls.cell(i_shop, 3).value)
                    print(sheets_xls.cell(i_shop, 4).value)

                    if isinstance(sheets_xls.cell(i_shop, 3).value, float):
                        last_year_value.append(sheets_xls.cell(i_shop, 3).value)
                    else:
                        last_year_value.append(0)
                    if isinstance (sheets_xls.cell(i_shop, 4).value, float):
                        current_year_value.append(sheets_xls.cell(i_shop, 4).value)
                    else:
                        current_year_value.append(0)
                    if isinstance(sheets_xls.cell(i_shop, 3).value, float) and isinstance(
                            sheets_xls.cell(i_shop, 4).value, float):
                        procent_year.append(
                            round((100 - (sheets_xls.cell(i_shop, 3).value / sheets_xls.cell(i_shop, 4).value * 100)),
                                  2))
                    else:
                        procent_year.append(0)
                else:
                    a1 = sheets_xls.cell(i_shop, (i * 4) - 1).value
                    a = sheets_xls.cell(i_shop, (i * 4)).value
                    if isinstance(a1, float):
                        last_year_value.append(sheets_xls.cell(i_shop, (i * 4) - 1).value)
                    else:
                        last_year_value.append(0)
                    if isinstance(a, float):
                        current_year_value.append(sheets_xls.cell(i_shop, (i * 4)).value)
                    else:
                        current_year_value.append(0)
                    if isinstance(a1, float) and isinstance(a, float):
                        procent_year.append(round(100 - (
                                    sheets_xls.cell(i_shop, (i * 4) - 1).value / sheets_xls.cell(i_shop, (
                                        i * 4)).value * 100), 2))
                    else:
                        procent_year.append(0)
            print(shop_list)
            print(last_year_value)
            print(current_year_value)
            print(procent_year)

        chart_data.categories = shop_list
        chart_data.add_series(last_year, (tuple(last_year_value)))
        chart_data.add_series(year, (tuple(current_year_value)))
        chart_data.add_series('%', (tuple(procent_year)))

        summ_last_year_weight.append(sum(last_year_value))
        summ_current_year_weight.append(sum(current_year_value))

        x, y, cx, cy = Inches(0.01), Inches(2), Inches(15), Inches(4.5)
        graphic_frame = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data)
        chart = graphic_frame.chart
        plot = chart.plots[0]
        plot.has_data_labels = True
        data_labels = plot.data_labels

        data_labels.font.size = Pt(13)
        data_labels.font.color.rgb = RGBColor(0x0A, 0x42, 0x80)
        data_labels.position = XL_LABEL_POSITION.INSIDE_END
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.RIGHT
        chart.legend.include_in_layout = False
        """
        Создание слайда с общими показателями за месяц
        """
        p.slide_width = Inches(15)
        ferst_slide = p.slide_layouts[5]

        # help(ferst_slide)
        slide = p.slides.add_slide(ferst_slide)
        s = p.slide_width
        print(s)
        slide.shapes.title.text = f'\n                Продажи к сравнению                  \nв киллограмах                               в рублях'

        slide.shapes.title.width = Inches(14)
        # slide.placeholders[1].text = f'Продажи к сравнению\n' \
        #                              f'{mount_title}'
        # define chart data ---------------------
        chart_data = ChartData()

        summ_procent_year_rub = []
        procent = round(100 - summ_last_year_rub[0] / summ_current_year_rub[0] * 100, 1)
        summ_procent_year_rub.append(procent)

        chart_data.categories = ['в рублях']
        chart_data.add_series(last_year, tuple(summ_last_year_rub))
        chart_data.add_series(year, tuple(summ_current_year_rub))
        chart_data.add_series('%', tuple(summ_procent_year_rub))



        x, y, cx, cy = Inches(1), Inches(2), Inches(6), Inches(4.5)
        graphic_frame = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data)
        chart = graphic_frame.chart
        plot = chart.plots[0]
        plot.has_data_labels = True
        data_labels = plot.data_labels

        data_labels.font.size = Pt(13)
        data_labels.font.color.rgb = RGBColor(0x0A, 0x42, 0x80)
        data_labels.position = XL_LABEL_POSITION.INSIDE_END
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.RIGHT
        chart.legend.include_in_layout = False

        chart_data_weight = ChartData()

        summ_procent_year_weight = []
        procent = round(100 - summ_last_year_weight[0] / summ_current_year_weight[0] * 100, 1)
        summ_procent_year_weight.append(procent)

        chart_data_weight.categories = ['В килkограмах']
        chart_data_weight.add_series(last_year, tuple(summ_last_year_weight))
        chart_data_weight.add_series(year, tuple(summ_current_year_weight))
        chart_data_weight.add_series('%', tuple(summ_procent_year_weight))

        x, y, cx, cy = Inches(9), Inches(2), Inches(6), Inches(4.5)
        graphic_frame = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data_weight)
        chart = graphic_frame.chart
        plot = chart.plots[0]
        plot.has_data_labels = True
        data_labels = plot.data_labels

        data_labels.font.size = Pt(13)
        data_labels.font.color.rgb = RGBColor(0x0A, 0x42, 0x80)
        data_labels.position = XL_LABEL_POSITION.INSIDE_END
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.RIGHT
        chart.legend.include_in_layout = False

        p.save('chart-01.pptx')



presentation_pptx_ferst(a)
presentation_pptx_shops_graphic(a, b)
    #             slide.add_empty_slide(pres.layout_slides[i])
    #             # Access first slide
    #             sld = pres.slides[i]
    #
    #             # Add chart with default data
    #             chart = sld.shapes.add_chart(slides.charts.ChartType.CLUSTERED_COLUMN, 0, 0, 500, 500)
    #
    #             # Set chart Title
    #             if i == 1:
    #                 mount_title = sheets_xls.cell(1, 1).value.split()
    #             else:
    #                 mount_title = sheets_xls.cell(1, (i * 4)-1).value.split()
    # create presentation with 1 slide ------



# for shape in slide.shapes:
#     if shape.is_placeholder:
#         phf = shape.placeholder_format
#         print('%d, %s' % (phf.idx, phf.type))





# def presentation_pptx_shops_graphic(file_name = ''):
#     file_xls = xlrd.open_workbook(file_name)
#     sheets_xls = file_xls.sheet_by_name("Анализ год к году")
#     colum_sh = sheets_xls.ncols
#     row_sh = sheets_xls.nrows
#     print(range(int(colum_sh - 1)//4))
#     with slides.Presentation('Анализ продаж по кварталам.pptx') as pres:
#         for i in range(1, int((colum_sh - 1)//4)-1):
#             print(i)
#              # Access first slide
#             slide = pres.slides
#             slide.add_empty_slide(pres.layout_slides[i])
#             # Access first slide
#             sld = pres.slides[i]
#
#             # Add chart with default data
#             chart = sld.shapes.add_chart(slides.charts.ChartType.CLUSTERED_COLUMN, 0, 0, 500, 500)
#
#             # Set chart Title
#             if i == 1:
#                 mount_title = sheets_xls.cell(1, 1).value.split()
#             else:
#                 mount_title = sheets_xls.cell(1, (i * 4)-1).value.split()
#             mount_title = mount_title[0]
#             chart.chart_title.add_text_frame_for_overriding(mount_title)
#             #chart.chart_title.text_frame_for_overriding.text_frame_format.center_text = 1
#             chart.chart_title.height = 20
#             chart.has_title = True
#
#             # Set first series to Show Values
#             chart.chart_data.series[0].labels.default_data_label_format.show_value = True
#
#             # Set the index of chart data sheet
#             defaultWorksheetIndex = 0
#
#             # Get the chart data worksheet
#             fact = chart.chart_data.chart_data_workbook
#
#             # Delete default generated series and categories
#             chart.chart_data.series.clear()
#             chart.chart_data.categories.clear()
#             s = len(chart.chart_data.series)
#             s = len(chart.chart_data.categories)
#
#             # Add new series
#             chart.chart_data.series.add(fact.get_cell(defaultWorksheetIndex, 0, 1, "Series 1"), chart.type)
#             chart.chart_data.series.add(fact.get_cell(defaultWorksheetIndex, 0, 2, "Series 2"), chart.type)
#
#             # Add new categories
#             chart.chart_data.categories.add(fact.get_cell(defaultWorksheetIndex, 1, 0, "Caetegoty 1"))
#             chart.chart_data.categories.add(fact.get_cell(defaultWorksheetIndex, 2, 0, "Caetegoty 2"))
#             chart.chart_data.categories.add(fact.get_cell(defaultWorksheetIndex, 3, 0, "Caetegoty 3"))
#
#             # Take first chart series
#             series = chart.chart_data.series[0]
#
#             # Now populating series data
#
#             series.data_points.add_data_point_for_bar_series(fact.get_cell(defaultWorksheetIndex, 1, 1, 20))
#             series.data_points.add_data_point_for_bar_series(fact.get_cell(defaultWorksheetIndex, 2, 1, 50))
#             series.data_points.add_data_point_for_bar_series(fact.get_cell(defaultWorksheetIndex, 3, 1, 30))
#
#             # Set fill color for series
#             series.format.fill.fill_type = slides.FillType.SOLID
#             series.format.fill.solid_fill_color.color = drawing.Color.red
#
#             # Take second chart series
#             series = chart.chart_data.series[1]
#
#             # Now populating series data
#             series.data_points.add_data_point_for_bar_series(fact.get_cell(defaultWorksheetIndex, 1, 2, 30))
#             series.data_points.add_data_point_for_bar_series(fact.get_cell(defaultWorksheetIndex, 2, 2, 10))
#             series.data_points.add_data_point_for_bar_series(fact.get_cell(defaultWorksheetIndex, 3, 2, 60))
#
#             # Setting fill color for series
#             series.format.fill.fill_type = slides.FillType.SOLID
#             series.format.fill.solid_fill_color.color = drawing.Color.orange
#
#             # First label will be show Category name
#             lbl = series.data_points[0].label
#             lbl.data_label_format.show_category_name = True
#
#             lbl = series.data_points[1].label
#             lbl.data_label_format.show_series_name = True
#
#             # Show value for third label
#             lbl = series.data_points[2].label
#             lbl.data_label_format.show_value = True
#             lbl.data_label_format.show_series_name = True
#             lbl.data_label_format.separator = "/"
#
#             # Save the presentation
#             pres.save("Анализ продаж по кварталам 11.pptx", slides.export.SaveFormat.PPTX)

