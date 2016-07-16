import xlsxwriter
from os import startfile

class excelSheet():
	def __init__(self):
		self.xmlFilePathCell ='A1'
		self.nameC = 'B'
		self.refC = 'C'
		self.typeC = 'D'
		self.titleRow = '3'
		self.firstInputRow = str(int(self.titleRow) + 1)
		self.wireTagCell = 'F3'
		self.wireCountCell = 'F4'
		self.GapsTagCell = 'A' + self.firstInputRow

	def startNewExcelSheet(self, xmlFilePath, xmlFolderPath, xmlFileName, refNameList, refGap, wireList):
		maxPossibleRefNum = 0
		xlsxFileName = xmlFileName + '_instruction.xlsx'
		xmlFilePath = xmlFilePath
		xlsxFilePath = xmlFolderPath + '/' + xlsxFileName
		workbook = xlsxwriter.Workbook(xlsxFilePath)
		worksheet = workbook.add_worksheet()
		
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

		worksheet.set_column(self.nameC + ':' + self.nameC, 10)
		worksheet.set_column(self.refC + ':' + self.refC, 25, unlocked)
		worksheet.set_column(self.typeC + ':' + self.typeC, 15, unlocked)
		worksheet.set_column(self.wireTagCell[0] + ':' + self.wireTagCell[0], 10, center)
		worksheet.set_row(0, None, locked)
		worksheet.set_row(1, None, locked)
		worksheet.set_row(2, None, locked)

		# Write a tags with different formats
		worksheet.write(self.xmlFilePathCell, 'Xml File Path: ' + xmlFilePath)
		worksheet.write(self.nameC + self.titleRow, "New Name", leftBottomBorder)
		worksheet.write(self.refC + self.titleRow, "Reference to Copy (R" + refNameList[0] + " - R" + refNameList[-1] + ")", bottomBorder)
		worksheet.write(self.typeC + self.titleRow, "Reference Type", bottomBorder)
		worksheet.write(self.wireTagCell, "Wire Count")
		worksheet.write(self.wireCountCell, len(wireList))

		### write gap reference names if there are gaps
		if refGap:
			lastGapRow = writeGapsName(worksheet, refGap, center, bottomBorder, border, self.nameC, self.GapsTagCell[0], int(self.firstInputRow))
		else:
			lastGapRow = self.titleRow

		### calculate the largest possible row
		maxPossibleRefNum = int(refNameList[-1]) - int(refNameList[0]) + 1 - (int(lastGapRow) - int(self.titleRow))
		fstAppendRow = str(int(lastGapRow) + 1)
		lstAppendRow = str(int(lastGapRow)+maxPossibleRefNum)

		if lastGapRow != self.titleRow:
			worksheet.write_blank(self.typeC+lastGapRow, None, unlockedBottomBorder)
			worksheet.write_blank(self.refC+lastGapRow, None, unlockedBottomBorder)

		### write reference name to append
		worksheet.merge_range(self.GapsTagCell[0] + fstAppendRow + ':' + self.GapsTagCell[0] + lstAppendRow, 'Append', right)
		writeAppendName(worksheet, int(refNameList[-1]) + 1 , fstAppendRow, lstAppendRow, self.nameC, center)

		### write comment
		
		worksheet.write_comment(self.refC + self.titleRow, 'Input any name of a existing refernce in the XML file (number only)')
		worksheet.write_comment(self.GapsTagCell, 'Must fill out all gaps')
		worksheet.write_comment(self.GapsTagCell[0] + fstAppendRow, 'This section can be empty')
		
		### set data validation
		worksheet.data_validation(self.refC + self.firstInputRow + ':' + self.refC + lstAppendRow, {'validate': 'list',           ### check double in the end
	                                 											     		   'source': refNameList,
	                                 											     		   'error_title': 'Warning',
			                                 											       'error_message': 'Reference does not exist!',
			                                 											       'error_type': 'stop'} )
		
		worksheet.data_validation(self.nameC + fstAppendRow + ':' + self.nameC + lstAppendRow, {'validate': 'integer',          
			                                 											        'criteria': '>',
			                                 											        'value': int(refNameList[-1]),
			                                 											        'error_title': 'Warning',
			                                 											        'error_message': 'Reference number too small or format incorrect!',
			                                 											        'error_type': 'stop'} )
		workbook.close()
		startfile(xlsxFilePath)

	def readExcelSheet(self):
		print(haha)






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




	

	












