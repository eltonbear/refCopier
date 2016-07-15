import xlsxwriter
from os import startfile

def startNewExcelSheet(xmlFilePath, xmlFolderPath, xmlFileName, refNameList, refGap, wireList):
	xlsxFileName = xmlFileName + '_instruction.xlsx'
	xlsxFilePath = xmlFolderPath + '/' + xlsxFileName
	workbook = xlsxwriter.Workbook(xlsxFilePath)
	worksheet = workbook.add_worksheet()


	# Create some cell formats with protection properties.
	unlocked = workbook.add_format({'locked': 0})
	locked = workbook.add_format({'locked': 1})
	hidden = workbook.add_format({'hidden': 1})
	center = workbook.add_format({'valign': 'vcenter', 'align': 'center'})

	# Format the columns to make the text more visible and editable
	worksheet.set_column('B:B', 25, unlocked)
	worksheet.set_column('C:C', 10)
	worksheet.set_column('D:D', 15, unlocked)
	worksheet.set_row(0, None, locked)
	worksheet.set_row(1, None, locked)
	worksheet.set_row(2, None, locked)

	# Turn worksheet protection on.
	worksheet.protect('elton')

	# Write a locked, unlocked and hidden cell.
	worksheet.write('A1', 'Xml File Path: ' + xmlFilePath)
	worksheet.write('B3', "Reference to Copy (R" + refNameList[0] + " - R" + refNameList[-1] + ")", center)
	worksheet.write('C3', "New Name", center)
	worksheet.write('D3', "Reference Type", center)
	worksheet.write('A4', "Gaps", center)
	worksheet.write('F3', "Wire Count", center)
	worksheet.write('F4', len(wireList), center)

	writeGaps(worksheet, refGap, center)

	workbook.close()
	startfile(xlsxFilePath)

def writeGaps(sheet, referenceGap, formatt):
	row = 4
	firstRow = row
	for gap in referenceGap:
		for name in range(gap[0] ,gap[-1]+1):
			sheet.write('C'+ str(row), "R" + str(name), formatt)
			row += 1
	sheet.merge_range('A'+ str(firstRow) + ':A' + str(row - 1), "Gaps", formatt)

def addRefToCopyRestriction(sheet, refNameList, referenceGap):
	print("nothing")

	






	

	












