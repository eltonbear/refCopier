import xlsxwriter
from tkinter import Tk
from os import startfile
import openpyxl as op
from warningWindow import errorMessage

class excelSheet():
	def __init__(self):
		self.xmlFilePathCell ='F1'
		self.nameC = 'B'
		self.refC = 'C'
		self.typeC = 'D'
		self.titleRow = '1'
		self.firstInputRow = str(int(self.titleRow) + 1)
		self.wireTagCell = 'F3'
		self.wireCountCell = 'F4'
		self.gapTagCell = 'A' + self.firstInputRow
		self.workSheetName = 'Reference_to_copy'
		self.gapTag = 'Fill Gaps'
		self.appendTag = 'Append'

	def startNewExcelSheet(self, xmlFilePath, xmlFolderPath, xmlFileName, refNumList, refGap, wireList):
		xlsxFileName = xmlFileName + '_instruction.xlsx'
		xmlPath = xmlFilePath
		xlsxFilePath = xmlFolderPath + '/' + xlsxFileName
		workbook = xlsxwriter.Workbook(xlsxFilePath)
		worksheet = workbook.add_worksheet(self.workSheetName)
		
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
		worksheet.set_column(self.wireTagCell[0] + ':' + self.wireTagCell[0], 10)

		# Write a tags with different formats
		worksheet.write(self.xmlFilePathCell, 'XML: ' + xmlPath)
		worksheet.write(self.nameC + self.titleRow, "New Name", leftBottomBorder)
		worksheet.write(self.refC + self.titleRow, "Reference to Copy (R" + refNumList[0] + " - R" + refNumList[-1] + ")", bottomBorder)
		worksheet.write(self.typeC + self.titleRow, "Reference Type", bottomBorder)
		worksheet.write(self.wireTagCell, "Wire Count", center)
		worksheet.write(self.wireCountCell, len(wireList), center)

		### write gap reference names if there are gaps
		if refGap:
			lastGapRow = str(int(self.titleRow) + len(refGap))
			fstAppendRow = str(int(lastGapRow) + 1)
			lstAppendRow = str(int(lastGapRow) +int(refNumList[-1]) - len(refGap))
			writeGapsName(worksheet, refGap, center, bottomBorder, border, self.nameC, self.gapTagCell[0], int(self.firstInputRow), self.gapTag)
			### add thicker bottom border
			worksheet.write_blank(self.typeC+lastGapRow, None, unlockedBottomBorder)
			worksheet.write_blank(self.refC+lastGapRow, None, unlockedBottomBorder)
		else:
			fstAppendRow = self.firstInputRow
			lstAppendRow = str(int(self.titleRow) + len(refNumList))

		### write reference name to append
		worksheet.merge_range(self.gapTagCell[0] + fstAppendRow + ':' + self.gapTagCell[0] + lstAppendRow, self.appendTag, right)
		writeAppendName(worksheet, int(refNumList[-1]) + 1 , fstAppendRow, lstAppendRow, self.nameC, center)

		### protect rows after the append section
		worksheet.set_row(int(lstAppendRow), None, locked)
		worksheet.set_row(int(lstAppendRow) + 1, None, locked)
		worksheet.set_row(int(lstAppendRow) + 2, None, locked)

		### write comment
		worksheet.write_comment(self.refC + self.titleRow, 'Input any name of a existing refernce in the XML file (number only)')
		worksheet.write_comment(self.gapTagCell, 'Must fill out all gaps')
		worksheet.write_comment(self.gapTagCell[0] + fstAppendRow, 'This section can be empty')
		
		### set data validation
		worksheet.data_validation(self.refC + self.firstInputRow + ':' + self.refC + lstAppendRow, {'validate': 'list',           ### check double in the end
	                                 											     		   'source': refNumList,
	                                 											     		   'error_title': 'Warning',
			                                 											       'error_message': 'Reference does not exist!',
			                                 											       'error_type': 'stop'} )
		
		# worksheet.data_validation(self.nameC + fstAppendRow + ':' + self.nameC + lstAppendRow, {'validate': 'integer',          
		# 	                                 											        'criteria': '>',
		# 	                                 											        'value': int(refNumList[-1]),
		# 	                                 											        'error_title': 'Warning',
		# 	                                 											        'error_message': 'Reference number too small or format incorrect!',
		# 	                                 											        'error_type': 'stop'} )
		
		try:
			workbook.close()
		except PermissionError:
			message = "Please close the existing Excel Sheet!"
			closeFileWindow = Tk()
			warning = errorMessage(closeFileWindow, message, None, False)
			closeFileWindow.mainloop()

		startfile(xlsxFilePath)
		
	def readExcelSheet(self, xlsxFilePath):
		excelName = []
		excelRef = []
		excelType = []
		workbook = op.load_workbook(filename = xlsxFilePath, read_only=True)
		try:
			worksheet = workbook.get_sheet_by_name(self.workSheetName)
		except keyError:
			message = "Cannot find excel sheet!"
			cantFindSheetWindow = Tk()
			warning = errorMessage(cantFindSheetWindow, message, None, False)
			cantFindSheetWindow.mainloop()
			return None
		
		xmlFilePath = worksheet[self.xmlFilePathCell].value[5:]

		row  = self.firstInputRow
		tagCo = worksheet.columns[0]
		lastRow = len(tagCo)
		excelName = []
		excelRef = []
		excelType = []
		missingRef = []
		missingType = []
		wrongSeqRow = []
		checkRepeatRef = set()
		temp = {}
		repeat = {}
		gap = tagCo[int(row)].value == self.gapTag
		prevBothNotEmpty = True
		while int(row) <= lastRow: 
			name = worksheet[self.nameC + row].value
			ref = str(worksheet[self.refC + row].value)
			typ = worksheet[self.typeC + row].value
			### gap
			if gap:
				if ref and typ and ref != 'None' and typ != 'None':
					excelName.append(name)
					excelRef.append(ref)
					excelType.append(typ)
				else:
					if not ref or ref == 'None':
						missingRef.append(row)
					if not typ or typ == 'None':
						missingType.append(row)
				if tagCo[int(row)].value == self.appendTag:
					gap = False
			else: ### append
				if ref and typ and  ref != 'None' and typ != 'None' and prevBothNotEmpty:
					excelName.append(name)
					excelRef.append(ref)
					excelType.append(typ)
				elif (not ref or ref == 'None') and typ and typ != 'None' and prevBothNotEmpty:
					missingRef.append(row)              
				elif (not typ or typ != 'None') and ref and ref != 'None' and prevBothNotEmpty:
					missingType.append(row)
				elif prevBothNotEmpty:
					prevBothNotEmpty = False
				elif (ref and ref != 'None') or (typ and typ != 'None'):
					wrongSeqRow.append(row)
					prevBothNotEmpty = True
					if not (ref and ref != 'None'):
						missingRef.append(row)
					elif not (typ and typ != 'None'):
						missingType.append(row)

			### chcek repeats
			if ref != 'None' and ref:
				if ref in temp and not ref in repeat:
					repeat[ref] = temp[ref]
					repeat[ref].append(row)
				elif ref in temp and ref in repeat:
					repeat[ref].append(row)
				else:
					temp[ref] = [row]

			row = str(int(row) + 1)

		# print(excelName)
		# print(excelRef)
		# print(excelType)		
		# print(missingRef)
		# print(missingType)
		# print(repeat)
		# print(wrongSeqRow)
		errorMessage = None
		if missingRef or missingType or repeat or wrongSeqRow:
			errorMessage = writeErrorMessage(missingRef, missingType, repeat, wrongSeqRow)
			print(errorMessage)
		return xmlFilePath, excelName, excelRef, excelType, errorMessage

def writeGapsName(sheet, referenceGap, format1, format2, format3, nameColumn, gapTagColumn, startingRow, gapTag):
	row = startingRow
	for missing in referenceGap:
		if missing == referenceGap[-1]:
			sheet.write(nameColumn + str(row), str(missing), format2)
		else:
			sheet.write(nameColumn + str(row), str(missing), format1)
		row += 1
	sheet.merge_range(gapTagColumn + str(startingRow) + ':' + gapTagColumn + str(row - 1), gapTag, format3)

	return str(row - 1)

def writeAppendName(sheet, refNameToAppend, firstAppendRow, lastAppendRow, nameColumn, format1):
	for row in range(int(firstAppendRow) , int(lastAppendRow) + 1):
		sheet.write(nameColumn + str(row), str(refNameToAppend), format1)
		refNameToAppend +=1

def writeErrorMessage(missingRefRow, missingTypeRow, repeatRefRow, wrongSequenceRow):
	message = ""
	if missingRefRow:
		message = message + "Missing Reference Number at Row: "
		for i in range(0, len(missingRefRow) - 1):
			message = message + missingRefRow[i] + ", "
		message = message + missingRefRow[-1] + "\n\n"

	if missingTypeRow:
		message = message + "Missing Reference Type at Row: "
		for i in range(0, len(missingTypeRow) - 1):
			message = message + missingTypeRow[i] + ", "
		message = message + missingTypeRow[-1] + "\n\n"

	if repeatRefRow:
		for ref in sorted(repeatRefRow.keys()):
			message = message + "R" + ref + " is repeated at row: "
			for i in range(0, len(repeatRefRow[ref]) -1):
				message = message + repeatRefRow[ref][i] + ", "
			message = message + repeatRefRow[ref][-1] + "\n"
		message = message + "\n"

	if wrongSequenceRow:
		message = message + "Sequence incorrect at row: "
		for i in range(0, len(wrongSequenceRow) - 1):
			message = message + wrongSequenceRow[i] + ", "
		message = message + wrongSequenceRow[-1] + "\n" 

	return message
	
# test = excelSheet()
# test.readExcelSheet(r"C:\Users\eltoshon\Desktop\programTestiing\xmltest1_instruction.xlsx")