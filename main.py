from tkinter import Tk
from firstInterface import first
from browseInterface import browse

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
		
	elif firstW.importSheet:
		### It's xlsx file
		window2 = Tk()
		importS= browse(window2, False)
		window2.mainloop()
		if importS.clickedBack:
			main()

main()

