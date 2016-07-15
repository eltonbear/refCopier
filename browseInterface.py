from tkinter import *
from tkinter.filedialog import askopenfilename
from os.path import isfile, split, splitext

class browse(Frame):
	def __init__(self, parent, isXML):
		Frame.__init__(self, parent, width = 1000)
		self.parent = parent
		self.filePath = ""
		self.folderPath = ""
		self.fileName = "" # with no extension
		self.filePathEntry = None
		self.isOk = False
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
		### Create buttons for Cancel, Ok, and Browse and set their positions
		if self.isXmlNotXlsx:
			browseText = "Browse xml"
		else:
			browseText = "Browse xlsx"
		bBrowse = Button(self.entryFrame, text = browseText, width = 12, command = self.getFilePath)	
		bBrowse.grid(row = 0, column = 1, padx=3, pady=3)
		bCancel = Button(self, text = "Cancel", width = 10 ,command = self.closeWindow)
		bCancel.pack(side = RIGHT,padx=5, pady=2)
		bOk = Button(self, text = "Ok", width = 5, command = self.OK)
		bOk.pack(side = RIGHT, padx=5, pady=2)

	def getFilePath(self):
		if self.isXmlNotXlsx:
			fileType = ("XML file", "*.xml")
		else:
			fileType = ("Excel Worksheet", "*.xlsx")
		self.filePath = askopenfilename(filetypes = (fileType, ("All files", "*.*")))
		self.filePathEntry.delete(0, 'end')
		self.filePathEntry.insert(0, self.filePath)

	def closeWindow(self):
		self.parent.destroy()

	def OK(self):
		### Command when Ok button is clicked	
		self.filePath = self.filePathEntry.get()						
		if self.filePath == "":
			self.emptyFileNameWarning()
		elif not isfile(self.filePath):
			self.incorrectFileNameWarning()
		else:
			self.isOk = True
			self.getFolderAndFileName()	
			self.closeWindow()

	def getFolderAndFileName(self):
		self.folderPath, self.fileName = split(self.filePath)
		self.fileName = splitext(self.fileName)[0]

	def incorrectFileNameWarning(self):
		messagebox.showinfo("Warning", "File does not exist!")

	def emptyFileNameWarning(self):
		messagebox.showinfo("Warning", "No files selected!")

	def fileFormatIncorrectWarning(self):
		messagebox.showinfo("Warning", "File: " + self.fileName + " - format incorrect!")