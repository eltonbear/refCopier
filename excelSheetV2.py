import xlsxwriter
from tkinter import Tk
from os import startfile
import openpyxl as op
from warningWindow import errorMessage
from util import splitFileFolderAndName

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
		self.lastAppendRowCell = 'F1'
		self.mTag = 'missing'
		self.eTag = 'existing'
		self.aTag = 'appending'
		self.workSheetName = 'Reference_copying'
		self.copyBlockedText = 'BLOCKED'

	def startNewExcelSheet(self, xmlFilePath, refNumList, refGap, typeList, depList, wireList):
		xmlFolderPath, xmlFileName = splitFileFolderAndName(xmlFilePath)
		xlsxFileName = xmlFileName + '_instruction.xlsx'
		xlsxFilePath = xmlFolderPath + '/' + xlsxFileName
		workbook = xlsxwriter.Workbook(xlsxFilePath)
		worksheet = workbook.add_worksheet(self.workSheetName)

		### add cell format
		unlocked = workbook.add_format({'locked': 0, 'valign': 'vcenter', 'align': 'center'})
		centerF = workbook.add_format({'valign': 'vcenter', 'align': 'center'})
		titleF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#b8cce0', 'font_color': '#1f497d', 'bold': True, 'bottom': 2, 'bottom_color': '#82a5d0'})
		copyBlockedF = workbook.add_format({'bg_color': '#a6a6a6', 'font_color': '#a6a6a6'})
		missingTagAndRefF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1, 'border_color': '#b2b2b2'   })
		missingUnblockedF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#FFC7CE', 'locked': 0, 'border': 1, 'border_color': '#b2b2b2'})
		missingDepBlockedF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#FFC7CE', 'locked': 1, 'hidden': 1,'border': 1, 'border_color': '#b2b2b2'})
		missingDepBlockedBlankF = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#FFC7CE', 'locked': 1, 'hidden': 1,'border': 1, 'border_color': '#b2b2b2'})
		existingWhiteBlockedF = workbook.add_format({'font_color': 'white', 'locked': 1})
		appendTagAndRefF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'font_color': 'white', 'bg_color': '#92cddc', 'locked': 1, 'border': 1, 'border_color': '#b2b2b2'})
		appendUnblockedF =  workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#92cddc', 'locked': 0,'border': 1, 'border_color': '#b2b2b2'})
		appendDepBlockedF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#92cddc', 'locked': 1, 'hidden': 1, 'border': 1, 'border_color': '#b2b2b2'})
		appendDepBlockedBlankF = workbook.add_format({'bg_color': '#92cddc', 'font_color': '#92cddc', 'locked': 1, 'hidden': 1, 'border': 1, 'border_color': '#b2b2b2'})

		### activate protection
		worksheet.protect('elton')

		### set column width and  protection
		worksheet.set_column(self.statusC + ':' + self.statusC, 10)
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
		worksheet.write(self.wireTagCell, "Wire Count", centerF)
		worksheet.write(self.wireCountCell, len(wireList), centerF)
		worksheet.write(self.xmlFilePathCell, 'XML: ' + xmlFilePath)

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
					worksheet.conditional_format(self.depC + rowS, {'type': 'cell', 'criteria': 'equal to', 'value': 0, 'format': missingDepBlockedBlankF})
					worksheet.data_validation(self.copyC + rowS, {'validate': 'list', 'source': refNumList,'error_title': 'Warning', 'error_message': 'Reference does not exist!', 'error_type': 'stop'}) 
				else:  ### existing ref row
					worksheet.write(self.statusC + rowS, self.eTag, existingWhiteBlockedF)
					worksheet.write(self.refC + rowS, refNumber, centerF)
					worksheet.write(self.copyC + rowS, self.copyBlockedText, copyBlockedF)
					worksheet.write(self.typeC + rowS, typeList[refListIndex],  unlocked)
					worksheet.write(self.depC + rowS, depList[refListIndex],  centerF)
					worksheet.conditional_format(self.depC + rowS, {'type': 'text', 'criteria': 'containing', 'value': 'none','format': existingWhiteBlockedF})
					worksheet.data_validation(self.depC + rowS, {'validate': 'list', 'source': refNumList,'error_title': 'Warning', 'error_message': 'Reference does not exist!', 'error_type': 'stop'})
					refListIndex += 1
			else: ### append section
				worksheet.write(self.statusC + rowS, self.aTag, appendTagAndRefF)
				worksheet.write(self.refC + rowS, refNumber, appendTagAndRefF)
				worksheet.write(self.copyC + rowS, None, appendUnblockedF)
				worksheet.write(self.typeC + rowS, None,  appendUnblockedF)
				worksheet.write_formula(self.depC + rowS, '=' + self.copyC + rowS, appendDepBlockedF)
				worksheet.conditional_format(self.depC + rowS, {'type': 'cell', 'criteria': 'equal to', 'value': 0, 'format': appendDepBlockedBlankF})
				worksheet.data_validation(self.copyC + rowS, {'validate': 'list', 'source': refNumList,'error_title': 'Warning', 'error_message': 'Reference does not exist!', 'error_type': 'stop'})
			refNumber = refNumber + 1

		worksheet.write(self.lastAppendRowCell, rowN,  existingWhiteBlockedF)


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
		try:
			workbook = op.load_workbook(filename = xlsxFilePath, read_only = True, data_only=True)
			worksheet = workbook.get_sheet_by_name(self.workSheetName)
		except op.utils.exceptions.InvalidFileException:
			return 0, 0, 0
		except KeyError:
			message = "Cannot find excel sheet - " + self.workSheetName + "!"
			cantFindSheetWindow = Tk()
			warning = errorMessage(cantFindSheetWindow, message, None, False)
			cantFindSheetWindow.mainloop()
			return None, None, None
			
		xmlFilePath = worksheet[self.xmlFilePathCell].value[5:]
		lastRow = worksheet[self.lastAppendRowCell].value # str

		excelReference = {'og': {}, 'add': {}, 'newRefName': []}
		missingRef = []
		missingCopy = []
		missingType = []
		missingDep = []
		wrongSeqRow = []
		newRefName = []
		checkRepeatRef = set()
		allCopy = {}
		repeat = {}
		row  = self.firstInputRow
		prevAllExist = True
		error = False
		while int(row) <= int(lastRow): 
			status = worksheet[self.statusC + row].value
			ref = str(worksheet[self.refC + row].value)
			copy = str(worksheet[self.copyC + row].value)
			typ = worksheet[self.typeC + row].value
			dep = str(worksheet[self.depC + row].value)

			refExists = ref and ref != 'None'
			copyExists = copy and copy !='None'
			typeExists = typ and typ !='None'
			depExists = dep and dep != 'None' and dep != '0' ### with formula 
			depCellEmpty = dep == None or dep == 'None'      ### if gets modified by users and left empty
			if dep == 'None':
				dep = None
			if status == self.eTag: 
				if refExists and copyExists and typeExists and not error:
					excelReference['og'][ref] = [typ, dep]
				else:
					if not refExists:
						missingRef.append(row)
					if not typeExists:
						missingType.append(row)
					error = True
			elif status == self.mTag:
				if refExists and copyExists and typeExists and depExists and not error:
					excelReference['add'][ref] = [copy, typ]
					excelReference['newRefName'].append(ref)
				else:
					if not refExists:
						missingRef.append(row)
					if not copyExists:
						missingCopy.append(row)
					if not typeExists:
						missingType.append(row)
					if depCellEmpty:
						missingDep.append(row)
					error = True
			else: ### append
				if prevAllExist:					
					if refExists and copyExists and typeExists and depExists and not error:
						excelReference['add'][ref] = [copy, typ]
						excelReference['newRefName'].append(ref)
					elif refExists and not copyExists and not typeExists and not depExists:
						prevAllExist = False
					else:
						if not refExists:
							missingRef.append(row)
						if not copyExists:
							missingCopy.append(row)
						if not typeExists:
							missingType.append(row)
						if depCellEmpty:
							missingDep.append(row)
						error = True
				elif copyExists or typeExists or depExists:
					wrongSeqRow.append(row)
					if not refExists:
						missingRef.append(row)
					if not copyExists:
						missingCopy.append(row)
					if not typeExists:
						missingType.append(row)
					if depCellEmpty:
						missingDep.append(row) 
					prevAllExist = True
					error = True
			### chcek repeats
			if copy == self.copyBlockedText:
				copy = None
			if copy != 'None' and copy:
				if copy in allCopy and not copy in repeat:
					repeat[copy] = allCopy[copy]
					repeat[copy].append(row)
					error = True
				elif copy in allCopy and copy in repeat:
					repeat[copy].append(row)
				else:
					allCopy[copy] = [row]

			row = str(int(row) + 1)
		errorText = None
		if missingRef or missingCopy or missingType or missingDep or repeat or wrongSeqRow:
			errorText = writeErrorMessage(missingRef, missingCopy, missingType, missingDep, repeat, wrongSeqRow)

		return xmlFilePath, excelReference, errorText

def writeErrorMessage(missingRefRow, missingCopyRow, missingTypeRow, missingDepRow, repeatRefRow, wrongSequenceRow):
	message = ""
	if missingRefRow:
		message = message + "\nMissing Reference Number at Row: "
		for i in range(0, len(missingRefRow) - 1):
			message = message + missingRefRow[i] + ", "
		message = message + missingRefRow[-1] + "\n"

	if missingCopyRow:
		message = message + "\nMissing Copying Number at Row: "
		for i in range(0, len(missingCopyRow) - 1):
			message = message + missingCopyRow[i] + ", "
		message = message + missingCopyRow[-1] + "\n"

	if missingTypeRow:
		message = message + "\nMissing Reference Type at Row: "
		for i in range(0, len(missingTypeRow) - 1):
			message = message + missingTypeRow[i] + ", "
		message = message + missingTypeRow[-1] + "\n"

	if missingDepRow:
		message = message + "\nMissing Dependent Number at Row: "
		for i in range(0, len(missingDepRow) - 1):
			message = message + missingDepRow[i] + ", "
		message = message + missingDepRow[-1] + "\n"
 
	if repeatRefRow:
		for ref in sorted(repeatRefRow.keys()):
			message = message + "\nR" + ref + " is repeated at Row: "
			for i in range(0, len(repeatRefRow[ref]) -1):
				message = message + repeatRefRow[ref][i] + ", "
			message = message + repeatRefRow[ref][-1]
		message = message + "\n"

	if wrongSequenceRow:
		message = message + "\nSequence Incorrect at Row: "
		for i in range(0, len(wrongSequenceRow) - 1):
			message = message + wrongSequenceRow[i] + ", "
		message = message + wrongSequenceRow[-1] + "\n"  

	return message