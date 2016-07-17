import xml.etree.ElementTree as ET
from xml.etree.ElementTree import ElementTree, Element, SubElement
import re

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
				# if prevNum + 1 == currNum - 1:
				# 	refNameGap.append([prevNum + 1])
				# else:
				# 	refNameGap.append([prevNum + 1, currNum - 1])
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
		
def fileFormatIncorrectWarning(fileName):
	messagebox.showinfo("Warning", "File: " + fileName + " - format incorrect!") #### maybe make another interface











	

