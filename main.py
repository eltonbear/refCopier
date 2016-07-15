from tkinter import *
from firstInterface import first
from startNewInterface import startNew

window1 = Tk()
firstW = first(window1)
window1.mainloop()

if firstW.start:
	window2 = Tk()
	startN= startNew(window2)
	window2.mainloop()
	if startN.isOk:
		print("ahahha")

elif firstW.importSheet:
	print("yes")
