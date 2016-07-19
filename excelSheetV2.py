import xlsxwriter
from tkinter import Tk
from os import startfile
import openpyxl as op
from warningWindow import errorMessage

class excelSheet():
	def __init__(self):
		self.statusC = 'A'
		self.refC = 'B'
		self.copyC = 'C'
		self.typeC = 'D'
		self.depC = 'E'
		self.titleRow = '1'
		self.firstInputRow = str(int(self.titleRow) + 1)
		self.xmlFilePathCell ='G1'
		self.wireTagCell = 'G3'
		self.wireCountCell = 'G4'
		self.mTag = 'Missing'
		self.appendTag = 'Append'
		self.workSheetName = 'Reference_copying'

	def startNewExcelSheet(self, xmlFilePath, xmlFolderPath, xmlFileName, refNumList, refGap, typeList, depList, wireList):
		xlsxFileName = xmlFileName + '_instruction.xlsx'
		xmlPath = xmlFilePath
		xlsxFilePath = xmlFolderPath + '/' + xlsxFileName
		workbook = xlsxwriter.Workbook(xlsxFilePath)
		worksheet = workbook.add_worksheet(self.workSheetName)

		### add cell format
		unlocked = workbook.add_format({'locked': 0, 'valign': 'vcenter', 'align': 'center'})
		centerF = workbook.add_format({'valign': 'vcenter', 'align': 'center'})
		titleF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#b8cce0', 'font_color': '#1f497d', 'bold': True, 'bottom': 2, 'bottom_color': '#82a5d0'})
		copyBlockedF = workbook.add_format({'bg_color': '#a6a6a6'})
		missingTagAndRefF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1, 'border_color': '#b2b2b2'   })
		missingUnblockedF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#FFC7CE', 'locked': 0, 'border': 1, 'border_color': '#b2b2b2'})
		missingDepBlockedF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#FFC7CE', 'locked': 1, 'hidden': 1,'border': 1, 'border_color': '#b2b2b2'})
		appendTagAndRefF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'font_color': 'white', 'bg_color': '#92cddc', 'locked': 1, 'border': 1, 'border_color': '#b2b2b2'})
		appendDepBlockedF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#92cddc', 'locked': 1, 'hidden': 1, 'border': 1, 'border_color': '#b2b2b2'})
		appendUnblockedF =  workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#92cddc', 'locked': 0,'border': 1, 'border_color': '#b2b2b2'})

		### activate protection
		worksheet.protect('elton')

		### set column width and  protection
		worksheet.set_column(self.refC + ':' + self.refC, 20)
		worksheet.set_column(self.typeC + ':' + self.typeC, 15)
		worksheet.set_column(self.depC + ':' + self.depC, 17)
		worksheet.set_column(self.wireTagCell[0] + ':' + self.wireTagCell[0], 10)

		### write title
		worksheet.write(self.statusC + self.titleRow, 'Status', titleF)
		worksheet.write(self.refC + self.titleRow, 'Reference Number (R)', titleF)
		worksheet.write(self.copyC + self.titleRow, 'Copy (R)', titleF)
		worksheet.write(self.typeC + self.titleRow, 'Reference Type', titleF)
		worksheet.write(self.depC + self.titleRow, 'Dependent On (R)', titleF)

		### write rows
		lastRefRow = int(self.titleRow) + int(refNumList[-1])
		fstAppendRow = str(int(lastRefRow) + 1)
		lastAppendRow = str(int(lastRefRow) + len(refNumList) - len(refGap))
		refNumber = 1		
		refListIndex = 0
		for rowN in range(int(self.firstInputRow), int(lastAppendRow) + 1):
			rowS = str(rowN)
			if rowN < int(fstAppendRow):
				if str(refNumber) in refGap: ### missing ref row
					worksheet.write(self.statusC + rowS, self.mTag,  missingTagAndRefF)
					worksheet.write(self.refC + rowS, refNumber,  missingTagAndRefF)
					worksheet.write(self.copyC + rowS, None,  missingUnblockedF)
					worksheet.write(self.typeC + rowS, None,  missingUnblockedF)
					worksheet.write_formula(self.depC + rowS, '=' + self.copyC + rowS, missingDepBlockedF)
					worksheet.data_validation(self.copyC + rowS, {'validate': 'list', 'source': refNumList,'error_title': 'Warning', 'error_message': 'Reference does not exist!', 'error_type': 'stop'}) 
				else:  ### existing ref row
					worksheet.write(self.refC + rowS, refNumber, centerF)
					worksheet.write(self.copyC + rowS, None, copyBlockedF)
					worksheet.write(self.typeC + rowS, typeList[refListIndex],  unlocked)
					worksheet.write(self.depC + rowS, depList[refListIndex],  centerF)
					# print(refListIndex)
					refListIndex += 1
			else: ### append section
				worksheet.write(self.refC + rowS, refNumber, appendTagAndRefF)
				worksheet.write(self.copyC + rowS, None, appendUnblockedF)
				worksheet.write(self.typeC + rowS, None,  appendUnblockedF)
				worksheet.write_formula(self.depC + rowS, '=' + self.copyC + rowS, appendDepBlockedF)
				worksheet.data_validation(self.copyC + rowS, {'validate': 'list', 'source': refNumList,'error_title': 'Warning', 'error_message': 'Reference does not exist!', 'error_type': 'stop'})
			refNumber = refNumber + 1
		if fstAppendRow == lastAppendRow:
			worksheet.write(self.statusC + fstAppendRow, self.appendTag, appendTagAndRefF)
		else:
			worksheet.merge_range(self.statusC + fstAppendRow + ':' + self.statusC + lastAppendRow, self.appendTag, appendTagAndRefF)

		### add comment
		copyTitleComment = 'Input a name of any existing refernces from the XML file (number only).\nAll gaps need to be filled out'
		worksheet.write_comment(self.copyC + self.titleRow, copyTitleComment, {'author': 'Elton', 'width': 250, 'height': 50})
		worksheet.write_comment(self.statusC + fstAppendRow, 'Optional Section', {'author': 'Elton', 'width': 100, 'height': 15})
		if refGap:
			worksheet.write_comment(self.statusC + str(int(refGap[0]) + int(self.titleRow)), 'Reference gaps in xml file' , {'author': 'Elton', 'width': 140, 'height': 15})
		
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
		excelDependon = []
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

				
# test = excelSheet()
# f = r"C:\Users\eltoshon\Desktop\programTestiing\xmltest1.xml"
# test.startNewExcelSheet(f, r"C:\Users\eltoshon\Desktop\programTestiing", 'xmltest1',['1', '2', '4', '5', '9'], ['3', '6', '7', '8'], ['100','100','100','100','100'], None)