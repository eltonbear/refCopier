from tkinter import *
from io import open
from os import startfile

class errorMessage(Frame):
	def __init__(self, parent, message, textFilePath):
		self.parent = parent
		self.message = message
		self.textFilePath = textFilePath
		self.initGUI()

	def initGUI(self):
		self.buttonFrame = Frame(self.parent)
		self.buttonFrame.pack(fill = BOTH, expand = True)
		self.messageFrame = Frame(self.buttonFrame, borderwidth = 1)
		self.parent.title("Error Message")
		self.messageFrame.pack(fill = BOTH, expand = True)
		self.makeButtons()
		var = StringVar()
		label = Message(self.messageFrame, textvariable=var, relief=RAISED, width = 1000)
		var.set(self.message)
		label.pack(fill = BOTH, expand = True)

	def makeButtons(self):
		bSave = Button(self.buttonFrame, text = "Save", width = 5, command = self.writeToText)
		bSave.pack(side = RIGHT, padx=5, pady=2)
		bOk = Button(self.buttonFrame, text = "Ok", width = 5, command = self.parent.destroy)
		bOk.pack(side = RIGHT, padx=3, pady=2)

	def writeToText(self):
		file = open(self.textFilePath,'w')
		file.write(self.message)
		file.close()
		startfile(self.textFilePath)
		self.parent.destroy()


