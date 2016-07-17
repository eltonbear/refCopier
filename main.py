from tkinter import *
from firstInterface import first
from browseInterface import browse, splitFileFolderAndName
from xmlInfo import writeInfoText
from excelSheet import excelSheet
from warningWindow import errorMessage
import xmlReader
import xmlModifier
from io import open
from os import startfile

window1 = Tk()
firstW = first(window1)
window1.mainloop()

if firstW.start:
	### It's xml file
	window2 = Tk()
	startN= browse(window2, True)
	window2.mainloop()
	if startN.isOk:
		refNameList, refGap, wireList = xmlReader.readXML(startN.filePath, startN.fileName)
		refNameRepeats = xmlReader.checkRepeats(refNameList)
		print(refNameList)
		print(refGap)
		print(refNameRepeats)
		if refNameRepeats:
			### creat info files when there is a repeat                   														########################### call error message
			writeInfoText(startN.filePath, startN.folderPath, startN.fileName, refNameRepeats, refNameList, refGap, wireList)
		else:
			### write excel sheet
			excelWrite = excelSheet()
			excelWrite.startNewExcelSheet(startN.filePath, startN.folderPath, startN.fileName, refNameList, refGap, wireList)

elif firstW.importSheet:
	### It's xlsx file
	window2 = Tk()
	importS= browse(window2, False)
	window2.mainloop()
	if importS.isOk:
		excelRead = excelSheet()
		xmlPath, excelNam, excelRef, excelTyp, error = excelRead.readExcelSheet(importS.filePath)

		if error:
			#### call error messager
			errorWindow = Tk()
			warning = errorMessage(errorWindow, error)
			errorWindow.mainloop()
			if warning.isSaved:
				### write to text file
				errorFilePath = importS.folderPath+ '/' +  importS.fileName + '_error.txt'
				errorMessageFile = open(errorFilePath,'w')
				errorMessageFile.write(error)
				errorMessageFile.close()
				startfile(errorFilePath)
				# startfile(importS.filePath)
		else:
			### call xml modifier
			xmlFolder, xmlFileName = splitFileFolderAndName(xmlPath)
			refNameList, refGap, referenceList, wireList, tree = xmlReader.readXML(xmlPath, xmlFileName)           ##### merge xmlreader and modifer
			newXmlFilePath = xmlModifier.modifier(xmlFolder, xmlFileName, excelRef, excelNam, excelTyp, referenceList,  wireList, tree)
			startfile(newXmlFilePath)

