from tkinter import *


class first(Frame):
	def __init__(self, parent):
		self.parent = parent
		self.initGUI()
		self.start = False
		self.importSheet = False

	def initGUI(self):
		self.parent.title("Reference Copier")
		self.buttonFrame = Frame(self.parent, width = 270, borderwidth = 1) ######## window not wide enough!
		self.buttonFrame.pack(fill = BOTH, expand = True)
		self.makeButtons()

	def makeButtons(self):
		### Create buttons for Cancel, Ok, and Browse and set their positions
		bStart = Button(self.buttonFrame, text = "Start New", width = 10, command = self.startNew)
		bStart.pack(pady = 2)
		bImport = Button(self.buttonFrame, text = "Import Sheet", width = 15,command = self.importSheet)
		bImport.pack(pady = 2)
		bCancel = Button(self.buttonFrame, text = "Cancel", width = 10 ,command = self.closeWindow)
		bCancel.pack(pady = 2)	

	def closeWindow(self):
		self.parent.destroy()

	def startNew(self):
		self.start = True
		self.closeWindow()

	def importSheet(self):
		self.importSheet = True
		self.closeWindow()