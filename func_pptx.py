from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import CategoryChartData
from pptx.util import Inches
from openpyxl import load_workbook
import xlrd
a = 'анализ продаж.xlsx'
def presentation_pptx(file_name = ''):
    """Функция создает файл продажи и первый слайд берет года из файла exel"""
    file_xls = xlrd.open_workbook('анализ продаж.xls')
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
    slide.placeholders[1].text = f'Продажи к сравнению\n' \
                              f'{last_year} - {year}'
    presentation.save('example.pptx')



presentation_pptx()