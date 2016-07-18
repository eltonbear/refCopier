from tkinter import Tk
from firstInterface import first
from browseInterface import browse, splitFileFolderAndName
from excelSheet import excelSheet
from warningWindow import errorMessage
import xmlTool
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
		refNameList, refGap, referenceList, wireList, tree = xmlTool.readXML(startN.filePath, startN.fileName)
		refNameRepeats = xmlTool.checkRepeats(refNameList)
		print(refNameList)
		print(refGap)
		print(refNameRepeats)
		if refNameRepeats:
			### creat info files when there is a repeat                
			info = xmlTool.XMLInfo(startN.filePath, startN.fileName, refNameRepeats, refNameList, refGap, wireList)
			errorFilePath = startN.folderPath+ '/' +  startN.fileName + '_info.txt'
			infoWindow = Tk()
			warning = errorMessage(infoWindow, info, errorFilePath, True)
			infoWindow.mainloop()
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
			errorFilePath = importS.folderPath+ '/' +  importS.fileName + '_error.txt'
			errorWindow = Tk()
			warning = errorMessage(errorWindow, error, errorFilePath, True)
			errorWindow.mainloop()
		else:
			### call xml modifier
			xmlFolder, xmlFileName = splitFileFolderAndName(xmlPath)
			refNameList, refGap, referenceList, wireList, tree = xmlTool.readXML(xmlPath, xmlFileName)           						##### merge xmlreader and modifer
			newXmlFilePath = xmlTool.modifier(xmlFolder, xmlFileName, excelRef, excelNam, excelTyp, referenceList,  wireList, tree)
			startfile(newXmlFilePath)

