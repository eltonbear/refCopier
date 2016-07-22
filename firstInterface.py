from tkinter import *
from browseInterface import browse

class first(Frame):
	def __init__(self, parent):
		self.parent = parent
		self.initGUI()

	def initGUI(self):
		self.parent.title("Reference Copier")
		self.buttonFrame = Frame(self.parent, width = 270, borderwidth = 1)
		self.buttonFrame.pack(fill = BOTH, expand = True)
		self.makeButtons()

	def makeButtons(self):
		bCancel = Button(self.buttonFrame, text = "Cancel", width = 10 ,command = self.closeWindow)
		bCancel.pack(side = RIGHT, padx = 5, pady = 5)
		bImport = Button(self.buttonFrame, text = "Import Sheet", width = 15,command = self.importSheet)
		bImport.pack(side = RIGHT, padx = 5, pady = 5)
		bStart = Button(self.buttonFrame, text = "Start New", width = 10, command = self.startNew)
		bStart.pack(side = RIGHT, padx = 5, pady = 5)

	def closeWindow(self):
		self.parent.destroy()

	def hideWinodw(self):
		self.parent.withdraw()

	def showWindow(self):
		self.parent.deiconify()

	def startNew(self):
		self.hideWinodw()
		windowBX = Toplevel()
		startN = browse(windowBX, self, True)
		windowBX.mainloop()

	def importSheet(self):
		self.hideWinodw()
		windowBS = Toplevel()
		importS = browse(windowBS, self, True)
		windowBS.mainloop()

def main():
	window = Tk()
	firstW = first(window)
	window.mainloop()

main()