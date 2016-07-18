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
	dependon = []
	numOfwire = len(wireE)
	numOfRef = len(referenceE)

	### obtain gaps
	for i in range(0, numOfRef):
		numberS = re.findall('\d+', referenceE[i].find('Name').text)[0]
		depS = referenceE[i].find('Dependon')
		if depS != None:
			depS = re.findall('\d+', depS.text)[0]
		refName.append(numberS)
		dependon.append(depS)
		if i > 0:
			currNum = int(numberS)
			if currNum - prevNum > 1:
				for missing in range(prevNum + 1, currNum):
					refNameGap.append(str(missing))
			prevNum = currNum
		else:
			prevNum = int(numberS)
		
	return refName, refNameGap, dependon, referenceE, wireE, tree

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


def modifier(xmlFolderPath, xmlFileName, refList, nameList, typeList, referenceE, wireE, tree):
	### make a ElementTree object and find its root (highest node)   
	root = tree.getroot() 
	### make two lists of all reference elements(objects) and wire elements(objects)
	numOfRef = len(referenceE)
	numOfWire = len(wireE)
	for n in range(0, len(refList)):
		### Create reference(copy) with according names, types, and dependency
		### And insert them after the last referenceSystem if ref entry is not empty
		if refList[n]:
			for r in referenceE:
				if 'R' + refList[n] == r.find('Name').text:
					print("im here")
					copy = writeARefCopy(r, refList[n], nameList[n], typeList[n])
					root.insert(int(nameList[n])-1, copy) 							################## error caused by insert with random index??
					break

		### change wire's des		
		modifyWireDesRef(refList[n], nameList[n], wireE)
	### write to a new xml file
	newXmlFilePath = xmlFolderPath + "/" + xmlFileName + "_new.xml"
	tree.write(newXmlFilePath)
	return newXmlFilePath

def writeARefCopy(refToCopy, oldName, newName, typ): 
	### creat new referenceSystem node
	newRefEle = Element('ReferenceSystem')
	newNameEle = SubElement(newRefEle, 'Name')
	newNameEle.text = 'R' + newName
	newTypeEle = SubElement(newRefEle, 'Type')
	newTypeEle.text = typ
	newDepEle = SubElement(newRefEle, 'Dependon')
	newDepEle.text = 'R' + oldName
	### creat two nodes that refer to original points objects(elements)
	for p in refToCopy.findall('Point'):
		newRefEle.append(p)
	### formatting xml text so it prints nicly 
	indent(newRefEle, 1)
	### return the reference(address) of the ref element
	return newRefEle

def modifyWireDesRef(oldDes, newDes, wireElement):
	# print("old: " + oldDes)
	# print("new: " + newDes +'\n')
	for wire in wireElement:
		secBondDes = wire.findall('Bond')[1].find('Refsys')
		# print("wire's des: " + secBondDes.text)
		if secBondDes.text == 'R' + oldDes:
			secBondDes.text = 'R' + newDes
			# print("changed to: " + secBondDes.text)
		# else:
			# print("no changes")
		# print("\n")

### in-place prettyprint formatter found online --> http://effbot.org/zone/element-lib.htm#prettyprint 
def indent(elem, level=0):
    i = "\n" + level*"  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
        for elem in elem:
            indent(elem, level+1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i














	

