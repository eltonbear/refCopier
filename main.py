from tkinter import *
from firstInterface import first
from browseInterface import browse
from xmlInfo import writeInfoText
from startExcelSheet import excelSheet
import xmlReader


window1 = Tk()
firstW = first(window1)
window1.mainloop()

if firstW.start:
	### It's xml file
	window2 = Tk()
	startN= browse(window2, True)
	window2.mainloop()
	if startN.isOk:
		print("ahahha")
		refNameList, refGap, wireList = xmlReader.readXML(startN.filePath, startN.fileName)
		print(refNameList)
		print(refGap)
		refNameRepeats = xmlReader.checkRepeats(refNameList)
		print(refNameRepeats)
		if refNameRepeats:
			### creat info files when there is a repeat
			print("there is repeat")
			writeInfoText(startN.filePath, startN.folderPath, startN.fileName, refNameRepeats, refNameList, refGap, wireList)
		else:
			### write excel sheet
			print("no repaets")
			excel = excelSheet()
			excel.startNewExcelSheet(startN.filePath, startN.folderPath, startN.fileName, refNameList, refGap, wireList)


elif firstW.importSheet:
	### It's xlsx file
	print("yes")
	window2 = Tk()
	importS= browse(window2, False)
	window2.mainloop()
