import xlsxwriter
from os import startfile



f = r"C:\Users\eltoshon\Desktop\ExcelTesting\test.xlsm"
workbook = xlsxwriter.Workbook(f)
worksheet = workbook.add_worksheet()

workbook.add_vba_project('vbaProject.bin')

worksheet.insert_button('B3', {'macro': 'getXMLPath',
                               'caption': 'Press Me',
                               'width': 80,
                               'height': 30})
workbook.close()
startfile(f)


