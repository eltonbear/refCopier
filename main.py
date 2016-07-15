from tkinter import *
from firstInterface import first
from startNewInterface import startNew
from xmlInfo import writeInfoText
import xmlReader


window1 = Tk()
firstW = first(window1)
window1.mainloop()

if firstW.start:
	window2 = Tk()
	startN= startNew(window2)
	window2.mainloop()
	if startN.isOk:
		print("ahahha")
		refNameList, refGap, wireList = xmlReader.readXML(startN.xmlFilePath, startN.xmlFileName)
		print(refNameList)
		print(refGap)
		refNameRepeats = xmlReader.checkRepeats(refNameList)
		print(refNameRepeats)
		if refNameRepeats:
			### creat info files when there is a repeat
			print("there is repeat")
			writeInfoText(startN.xmlFilePath, startN.xmlFolderPath, startN.xmlFileName, refNameRepeats, refNameList, refGap, wireList)
		else:
			### write excel sheet
			print("no repaets")



elif firstW.importSheet:
	print("yes")
