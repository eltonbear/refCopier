from tkinter import *

class errorMessage(Frame):
	def __init__(self, parent, message):
		# Frame.__init__(self, parent, width = 500)
		self.parent = parent
		self.message = message
		self.isSaved = False
		self.isOk = False
		self.initGUI()

	def initGUI(self):
		self.buttonFrame = Frame(self.parent)
		self.buttonFrame.pack(fill = BOTH, expand = True)
		self.messageFrame = Frame(self.buttonFrame, borderwidth = 1)
		self.parent.title("Error Message")
		self.messageFrame.pack(fill = BOTH, expand = True)
		self.makeButtons()
		####
		var = StringVar()
		label = Message(self.messageFrame, textvariable=var, relief=RAISED, width = 550)
		var.set(self.message)
		label.pack(fill = BOTH, expand = True)

	def makeButtons(self):
		### Create buttons for Cancel, Ok, and Browse and set their positions
		bSave = Button(self.buttonFrame, text = "Save", width = 5, command = self.save)
		bSave.pack(side = RIGHT,padx=5, pady=2)
		bOk = Button(self.buttonFrame, text = "Ok", width = 5, command = self.OK)
		bOk.pack(side = RIGHT, padx=3, pady=2)

	def closeWindow(self):
		self.parent.destroy()

	def OK(self):
		self.isOk = True
		self.parent.destroy()

	def save(self):
		self.isSaved = True
		self.parent.destroy()

	

# a= Tk()
# x = errorMessage(a, "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa")
# a.mainloop()


