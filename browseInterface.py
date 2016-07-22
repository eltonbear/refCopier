from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
from os.path import isfile
from util import splitFileFolderAndName
from excelSheetV2 import excelSheet
from warningWindow import errorMessage
import xmlTool
from os import startfile


class browse(Frame):
	def __init__(self, parent, mainLevel, isXML):
		Frame.__init__(self, parent, width = 1000)
		self.parent = parent
		self.mainLevel = mainLevel
		self.filePath = ""
		self.filePathEntry = None
		self.isXmlNotXlsx = isXML
		self.initGUI()

	def initGUI(self):
		self.parent.title("Reference Copying")
		self.pack(fill = BOTH, expand = True)

		self.entryFrame = Frame(self, relief = RAISED, borderwidth = 1)
		self.entryFrame.pack(fill = BOTH, expand = True)
		self.makeButtons()

		self.filePathEntry = Entry(self.entryFrame, bd = 4, width = 50)
		self.filePathEntry.grid(row = 0, column = 2, columnspan = 5, padx=2, pady=2)

	def makeButtons(self):
		if self.isXmlNotXlsx:
			browseText = "Browse xml"
		else:
			browseText = "Browse xlsx"
		bBrowse = Button(self.entryFrame, text = browseText, width = 12, command = self.getFilePath)	
		bBrowse.grid(row = 0, column = 1, padx=3, pady=3)
		bCancel = Button(self, text = "Cancel", width = 10 ,command = self.closeMainAndToplevelWindow)
		bCancel.pack(side = RIGHT,padx=4, pady=2)
		bBack = Button(self, text = "Back", width = 7, command = self.back)
		bBack.pack(side = RIGHT,padx=4, pady=2)
		bOk = Button(self, text = "Ok", width = 5, command = self.OK)
		bOk.pack(side = RIGHT, padx=4, pady=2)

	def getFilePath(self):
		if self.isXmlNotXlsx:
			fileType = ("XML file", "*.xml")
		else:
			fileType = ("Excel Worksheet", "*.xlsx")
		self.filePath = askopenfilename(filetypes = (fileType, ("All files", "*.*")), parent = self.parent)
		self.filePathEntry.delete(0, 'end')
		self.filePathEntry.insert(0, self.filePath)

	def closeMainAndToplevelWindow(self):
		self.mainLevel.closeWindow()

	def closeWindow(self):
		self.parent.destroy()

	def OK(self):
		self.filePath = self.filePathEntry.get()						
		if self.filePath == "":
			self.emptyFileNameWarning()
		elif not isfile(self.filePath):
			self.incorrectFileNameWarning()
		else:
			folderPath, fileName = splitFileFolderAndName(self.filePath)
			if self.isXmlNotXlsx:
				result = readXMLAndStartSheet(self.filePath, folderPath, fileName)
				if type(result) == str:
					self.popErrorMessage(result)
				elif type(result) == list:
					### creat info files when there is a repeat  
					self.closeMainAndToplevelWindow()
					infoWindow = Tk()
					errorMessage(infoWindow, result[0], result[1])
					infoWindow.mainloop()
				else:
					self.closeMainAndToplevelWindow()
			else:
				result = readSheetAndModifyXML(self.filePath, folderPath, fileName)
				if type(result) == str:
					self.popErrorMessage(result)
				elif type(result) == list:
					self.closeMainAndToplevelWindow()
					errorWindow = Tk()
					errorMessage(errorWindow, result[0], result[1])
					errorWindow.mainloop()
				else:
					self.closeMainAndToplevelWindow()
			
	def back(self):
		self.closeWindow()
		self.mainLevel.showWindow()

	def incorrectFileNameWarning(self):
		messagebox.showinfo("Warning", "File does not exist!", parent = self.parent)

	def emptyFileNameWarning(self):
		messagebox.showinfo("Warning", "No files selected!", parent = self.parent)

	def popErrorMessage(self, message):
		messagebox.showinfo("Warning", message, parent = self.parent)

def readXMLAndStartSheet(filePath, folderPath, fileName):
	refNameList, refGap, typeList, depList, wireList = xmlTool.readXML(filePath)
	if not refNameList or not wireList:
		return "File: " + fileName + " - format incorrect!"
	refNameRepeats = xmlTool.checkRepeats(refNameList)
	if refNameRepeats:        
		info = xmlTool.XMLInfo(filePath, refNameRepeats, refNameList, refGap, wireList)
		errorFilePath = folderPath + '/' +  fileName + '_info.txt'
		return [info, errorFilePath]
	else:  
		excelWrite = excelSheet()
		error = excelWrite.startNewExcelSheet(filePath, refNameList, refGap, typeList, depList, wireList)
		return error
	
def readSheetAndModifyXML(filePath, folderPath, fileName):
	excelRead = excelSheet()
	xmlPath, refExcelDict, error = excelRead.readExcelSheet(filePath)
	print(xmlPath, refExcelDict, error)
	if xmlPath and refExcelDict:
		if error:
			### there is an error
			errorFilePath = folderPath + '/' + fileName + '_error.txt'
			return [error, errorFilePath]
		else:
			### call xml modifier
			newXmlFilePath = xmlTool.modifier(xmlPath, refExcelDict)
			if isfile(newXmlFilePath):
				startfile(newXmlFilePath)
				return None
			else:
				error = newXmlFilePath
				return error
	else:
		return error



