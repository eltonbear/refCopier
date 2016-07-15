import xlsxwriter
from os import startfile

def startNewExcelSheet(xmlFilePath, xmlFolderPath, xmlFileName, refNameList, refGap, wireList):
	xlsxFileName = xmlFileName + '_instruction.xlsx'
	xlsxFilePath = xmlFolderPath + '/' + xlsxFileName
	workbook = xlsxwriter.Workbook(xlsxFilePath)
	worksheet = workbook.add_worksheet()
	maxPossibleRefNum = 0 
	xmlFilePathCell ='A1'
	nameC = 'B'
	refC = 'C'
	typeC = 'D'
	titleRow = '3'
	firstInputRow = str(int(titleRow) + 1)
	wireTagCell = 'F3'
	wireCountCell = 'F4'
	GapsTagCell = 'A' + firstInputRow
	nextNewRow = '0'

	# Create some cell formats with protection properties.
	locked = workbook.add_format({'locked': 1})
	unlocked = workbook.add_format({'locked': 0, 'valign': 'vcenter', 'align': 'center'})
	center = workbook.add_format({'valign': 'vcenter', 'align': 'center'})
	border =  workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bottom': 5, 'top': 5, 'right': 5})
	right = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'right': 5, 'top': 5})
	leftBottomBorder = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bottom': 5, 'left': 5})
	bottomBorder = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bottom': 5})
	unlockedBottomBorder = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bottom': 5, 'locked': 0})

	# Turn worksheet protection on.
	worksheet.protect('elton')

	worksheet.set_column(nameC + ':' + nameC, 10)
	worksheet.set_column(refC + ':' + refC, 25, unlocked)
	worksheet.set_column(typeC + ':' + typeC, 15, unlocked)
	worksheet.set_column(wireTagCell[0] + ':' + wireTagCell[0], 10, center)
	worksheet.set_row(0, None, locked)
	worksheet.set_row(1, None, locked)
	worksheet.set_row(2, None, locked)

	# Write a tags with different formats
	worksheet.write(xmlFilePathCell, 'Xml File Path: ' + xmlFilePath)
	worksheet.write(nameC + titleRow, "New Name", leftBottomBorder)
	worksheet.write(refC + titleRow, "Reference to Copy (R" + refNameList[0] + " - R" + refNameList[-1] + ")", bottomBorder)
	worksheet.write(typeC + titleRow, "Reference Type", bottomBorder)
	worksheet.write(wireTagCell, "Wire Count")
	worksheet.write(wireCountCell, len(wireList))

	### write gap reference names if there are gaps
	if refGap:
		lastGapRow = writeGapsName(worksheet, refGap, center, bottomBorder, border, nameC, GapsTagCell[0], int(firstInputRow))
	else:
		lastGapRow = titleRow

	### calculate the largest possible row
	maxPossibleRefNum = int(refNameList[-1]) - int(refNameList[0]) + 1 - (int(lastGapRow) - int(titleRow))
	fstAppendRow = str(int(lastGapRow) + 1)
	lstAppendRow = str(int(lastGapRow)+maxPossibleRefNum)

	if lastGapRow != titleRow:
		worksheet.write_blank(typeC+lastGapRow, None, unlockedBottomBorder)
		worksheet.write_blank(refC+lastGapRow, None, unlockedBottomBorder)

	### write reference name to append
	worksheet.merge_range(GapsTagCell[0] + fstAppendRow + ':' + GapsTagCell[0] + lstAppendRow, 'Append', right)
	writeAppendName(worksheet, int(refNameList[-1]) + 1 , fstAppendRow, lstAppendRow, nameC, center)

	### write comment
	
	worksheet.write_comment(refC + titleRow, 'Input any name of a existing refernce in the XML file (number only)')
	worksheet.write_comment(GapsTagCell, 'Must fill out all gaps')
	worksheet.write_comment(GapsTagCell[0] + fstAppendRow, 'This section can be empty')
	
	### set data validation
	worksheet.data_validation(refC + firstInputRow + ':' + refC + lstAppendRow, {'validate': 'list',           ### check double in the end
                                 											   'source': refNameList,
                                 											   'error_title': 'Warning',
                                 											   'error_message': 'Reference does not exist!',
                                 											   'error_type': 'stop'} )
	
	worksheet.data_validation(nameC + fstAppendRow + ':' + nameC + lstAppendRow, {'validate': 'integer',          
                                 											   'criteria': '>',
                                 											   'value': int(refNameList[-1]),
                                 											   'error_title': 'Warning',
                                 											   'error_message': 'Reference number too small or format incorrect!',
                                 											   'error_type': 'stop'} )
	workbook.close()
	startfile(xlsxFilePath)

def writeGapsName(sheet, referenceGap, format1, format2, format3, nameColumn, gapTagColumn, startingRow):
	row = startingRow
	# missing = []
	for gap in referenceGap:
		for name in range(gap[0] ,gap[-1]+1):
			if gap == referenceGap[-1] and name == gap[-1]:
				sheet.write(nameColumn + str(row), str(name), format2)
			else:
				sheet.write(nameColumn + str(row), str(name), format1)
			# missing.append(name)
			row += 1
	sheet.merge_range(gapTagColumn + str(startingRow) + ':' + gapTagColumn + str(row - 1), 'Fill Gaps', format3)

	return str(row - 1)

def writeAppendName(sheet, refNameToAppend, firstAppendRow, lastAppendRow, nameColumn, format1):
	for row in range(int(firstAppendRow) , int(lastAppendRow) + 1):
		sheet.write(nameColumn + str(row), str(refNameToAppend), format1)
		refNameToAppend +=1



	

	












