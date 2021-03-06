from tkinter import Tk
from firstInterface import first
from browseInterface import browse
from excelSheetV2 import excelSheet
from warningWindow import errorMessage
from os import startfile
from util import splitFileFolderAndName
import xmlTool

def main():
	window1 = Tk()
	firstW = first(window1)
	window1.mainloop()
	if firstW.start:
		### It's xml file
		window2 = Tk()
		startN= browse(window2, True)
		window2.mainloop()
		if startN.clickedBack:
			main()
		elif startN.isOk:
			refNameList, refGap, typeList, depList, wireList = xmlTool.readXML(startN.filePath)
			folderPath, fileName = splitFileFolderAndName(startN.filePath)
			if not refNameList or not wireList:
				startN.fileFormatIncorrectWarning(fileName)
			else:
				refNameRepeats = xmlTool.checkRepeats(refNameList)
				if refNameRepeats:
					### creat info files when there is a repeat            
					info = xmlTool.XMLInfo(startN.filePath, refNameRepeats, refNameList, refGap, wireList)
					errorFilePath = folderPath + '/' +  fileName + '_info.txt'
					infoWindow = Tk()
					warning = errorMessage(infoWindow, info, errorFilePath, True)
					infoWindow.mainloop()
				else:
					### write excel sheet
					excelWrite = excelSheet()
					excelWrite.startNewExcelSheet(startN.filePath, refNameList, refGap, typeList, depList, wireList)
	elif firstW.importSheet:
		### It's xlsx file
		window2 = Tk()
		importS= browse(window2, False)
		window2.mainloop()
		if importS.clickedBack:
			main()
		elif importS.isOk:
			excelRead = excelSheet()
			xmlPath, refExcelDict, error = excelRead.readExcelSheet(importS.filePath)
			folderPath, fileName = splitFileFolderAndName(importS.filePath)
			if xmlPath == 0 and refExcelDict == 0 and error == 0:
				importS.fileFormatIncorrectWarning(fileName)
			elif xmlPath and refExcelDict:
				if error:
					#### call error messager
					errorFilePath = folderPath + '/' + fileName + '_error.txt'
					errorWindow = Tk()
					warning = errorMessage(errorWindow, error, errorFilePath, True)
					errorWindow.mainloop()
				else:
					### call xml modifier
					newXmlFilePath = xmlTool.modifier(xmlPath, refExcelDict)
					if not newXmlFilePath:
						importS.fileFormatIncorrectWarning(fileName)
					else:
						startfile(newXmlFilePath)

main()

