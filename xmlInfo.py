from io import open
from os import startfile
from os.path import exists

def writeInfoText(xmlFilePath, xmlFolderPath, xmlFileName, repRef, refName, refGap, wireList):

	infoName = xmlFileName + "_info.txt" 
	infoFilePath = xmlFolderPath + "/" + infoName
	numR = len(refName)
	numW = len(wireList)


	if exists(xmlFolderPath):
		info = open(infoFilePath, "w")
		### write file path		
		info.write("#Input XML File: " + xmlFilePath + '\n\n')
		### write repeating ref name if there is any
		info.write("#Repeating Reference:\n")
		if repRef:
			for r in repRef:
				info.write("There are " + str(r[1]) + " R" + r[0] + '\n')
		else:
			info.write("None\n")
		### write first and last ref name
		info.write("\n#First Reference: R" + refName[0] + '\n')
		info.write("#Last Reference:  R" + refName[numR-1] + '\n')
		### write refernce gaps
		info.write("\n#Range of Gaps (included):\n")
		if refGap:
			for g in refGap:
				if len(g) == 1:
					info.write("R" + str(g[0]) + '\n')
				else:
					info.write("R" + str(g[0]) + ' - R' + str(g[1]) + '\n')
		else:
			info.write("None\n")
		info.write("\n#Number of Wires: " + str(numW) + "\n")

		### Close file
		info.close()
		### open text file in windows
		startfile(infoFilePath)
	

