import xml.etree.ElementTree as ET
from xml.etree.ElementTree import ElementTree, Element, SubElement
import re
from os.path import exists

def readXML(xmlFilePath, xmlFileName):
	### check format
	try:
		tree = ET.parse(xmlFilePath)                                    
	except ET.ParseError:                       
		fileFormatIncorrectWarning(xmlFileName)
		return None

	root = tree.getroot() 
	referenceE = root.findall('ReferenceSystem')
	wireE = root.findall('Wire')
	### check format
	if not referenceE or not wireE:
		fileFormatIncorrectWarning()
		return None

	refName = [] #str
	refNameGap = [] #list of lists of int(single or pair)
	numOfwire = len(wireE)
	numOfRef = len(referenceE)

	### obtain gaps
	for i in range(0, numOfRef):
		numberS = re.findall('\d+', referenceE[i].find('Name').text)[0]
		refName.append(numberS)
		if i > 0:
			currNum = int(numberS)
			if currNum - prevNum > 1:
				for missing in range(prevNum + 1, currNum):
					refNameGap.append(str(missing))
			prevNum = currNum
		else:
			prevNum = int(numberS)
		
	return refName, refNameGap, referenceE, wireE, tree

def checkRepeats(refNameList):
	''' Check if there is any repeating reference. Return a list of lists of a name of repeating ref(str) and count(int)'''
	repeat = []
	singles = set(refNameList)
	for s in singles:
		count = refNameList.count(s)
		if count > 1:
			repeat.append([s, count])
	return repeat


def XMLInfo(xmlFilePath, xmlFileName, repRef, refName, refGap, wireList):
	numR = len(refName)
	numW = len(wireList)

	if exists(xmlFilePath):
		info = ""
		### write file path		
		info = info + "#Input XML File: " + xmlFilePath + '\n\n'
		### write repeating ref name if there is any
		info = info + "#Repeating Reference:\n"
		if repRef:
			for r in repRef:
				info = info + "There are " + str(r[1]) + " R" + r[0] + '\n'
		else:
			info  = info + "None\n"
		### write first and last ref name
		info = info + "\n#First Reference: R" + refName[0] + '\n'
		info = info + "#Last Reference:  R" + refName[numR-1] + '\n'
		### write refernce gaps
		info = info + "\n#Range of Gaps (included):\n"
		if refGap:
			for g in refGap:
				if len(g) == 1:
					info = info + "R" + str(g[0]) + '\n'
				else:
					info = info + "R" + str(g[0]) + ' - R' + str(g[1]) + '\n'
		else:
			info = info + "None\n"
		info = info + "\n#Number of Wires: " + str(numW) + "\n"
	else:
		info = "File does not exist!"

	return info
		
def fileFormatIncorrectWarning(fileName):
	messagebox.showinfo("Warning", "File: " + fileName + " - format incorrect!") #### maybe make another interface











	

