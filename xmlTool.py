import xml.etree.ElementTree as ET
from xml.etree.ElementTree import ElementTree, Element, SubElement
import re
from os.path import exists
from util import splitFileFolderAndName

def readXML(xmlFilePath):
	"""Read a XML data file.

		Parameters
		----------
		xmlFilePath: string

		Returns
		-------


	"""

	parseFailure = False
	try:
		# Parse xml content
		tree = ET.parse(xmlFilePath)                                    
	except ET.ParseError: 
		# True if it fails parsing
		parseFailure  = True
	# Get root of the xml data tree and all reference and wire elements if xml is parsed successfullyncorrect
	if not parseFailure:
		root = tree.getroot() 
		referenceE = root.findall('ReferenceSystem')
		wireE = root.findall('Wire')
	# if there is an error in parsing or getting elements, return None for all 
	if parseFailure or not referenceE or not wireE:
		return None, None, None, None, None

	refName = [] 				# A list of reference numbers in string
	refNameGap = [] 			# A list of missing reference numbers in string
	typ = [] 					# A list of refernece types in string
	dependon = [] 				# A list of reference number that references depend on in string
	numOfRef = len(referenceE)	# The number of reference elements

	# Obtain wire source and destination information and total count 
	wireSDCount = readWireSDInfo(wireE)
	
	# Obtain ref, type, dep, and gaps
	for i in range(0, numOfRef):
		numberS = re.findall('\d+', referenceE[i].find('Name').text)[0]
		depS = referenceE[i].find('Dependon')
		if depS != None:
			depS = re.findall('\d+', depS.text)[0]
		refName.append(numberS)
		typ.append(referenceE[i].find('Type').text)
		dependon.append(depS)
		if i > 0:
			currNum = int(numberS)
			if currNum - prevNum > 1:
				for missing in range(prevNum + 1, currNum):
					refNameGap.append(str(missing))
			prevNum = currNum
		else:
			prevNum = int(numberS)

		if not numberS in wireSDCount:
			wireSDCount[numberS] = {'s': [], 'd': []}
	
	return refName, refNameGap, typ, dependon, wireSDCount

def readWireSDInfo(wireElements):
	### obtain wire count --> wireSDInfo = {totalWireCount: n, 'refNum': {s:[wireIndex], d:[wireIndex]}}
	wireSDInfo = {'total': len(wireElements)}
	for wireIndex in range(0, len(wireElements)):
		source = re.findall('\d+', wireElements[wireIndex].findall('Bond')[0].find('Refsys').text)[0] #str
		destination = re.findall('\d+', wireElements[wireIndex].findall('Bond')[1].find('Refsys').text)[0] #str
		if source in wireSDInfo:
			wireSDInfo[source]['s'].append(wireIndex)
		else:
			wireSDInfo[source] = {'s': [wireIndex], 'd': []}
		if destination in wireSDInfo:
			wireSDInfo[destination]['d'].append(wireIndex)
		else:
			wireSDInfo[destination] = {'s': [], 'd': [wireIndex]}
	return wireSDInfo

def checkRepeats(refNameList):
	''' Check if there is any repeating reference. Return a list of lists of a name of repeating ref(str) and count(int)'''
	repeat = []
	singles = set(refNameList)
	for s in singles:
		count = refNameList.count(s)
		if count > 1:
			repeat.append([s, count])
	return repeat

def XMLInfo(xmlFilePath, repRef, refName, refGap, wireSDCount):
	numR = len(refName)
	numW = wireSDCount['total']
	prefix = 'R'

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
					info = info + prefix + str(g[0]) + '\n'
				else:
					info = info + prefix + str(g[0]) + ' - ' + prefix + str(g[1]) + '\n'
		else:
			info = info + "None\n"
		info = info + "\n#Number of Wires: " + str(numW) + "\n"
	else:
		info = "File does not exist!"

	return info
		
def modifier(xmlFilePath, referenceDictDFromExc):
	xmlFolderPath, xmlFileName = splitFileFolderAndName(xmlFilePath)
	### make a ElementTree object and find its root (highest node)
	### if it fails parsing, file format is incorrect  
	try:
		tree = ET.parse(xmlFilePath)                                    
	except ET.ParseError: 
		message = "File: " + xmlFileName + " - format incorrect!"
		return message

	root = tree.getroot() 
	### make two lists of all reference elements(objects) and wire elements(objects)
	referenceE = root.findall('ReferenceSystem')
	wireE = root.findall('Wire') 
	numOfRef = len(referenceE)
	numOfWire = len(wireE)
	prefix = 'R'
	for r in referenceE: ### Modify existing ref's type and dep if necessary
		ref = r.find('Name').text
		refNumber = re.findall('\d+', ref)[0]
		typ, dep = referenceDictDFromExc['og'][refNumber]
		if r.find('Type').text != typ:
			r.find('Type').text = typ
		if dep:
			if r.find('Dependon') == None:
				newDepEle = Element('Dependon')
				newDepEle.text = prefix + dep
				r.insert(2, newDepEle)
				indent(newDepEle, 2)
			elif r.find('Dependon').text != prefix + dep:
				r.find('Dependon').text = prefix + dep
		else:
			if r.find('Dependon') != None:
				r.remove(r.find('Dependon'))

	### Data Structure:
	### read wire source and destination information and return a dictionary --> wireSDInfo = {totalWireCount: n, 'refNum': {s:[wireIndex], d:[wireIndex]}}
	### referenceDictDFromExc data structure --> {'og': {'refNum':[type, dependon]}, 'add': {'refNum': [copyNum, type]}, 'newRefName': [str(refNum)]}
	### addRefDict --> {'str(refNum)': ['copyNum', type]}
	### referenceDictDFromExc['newRefName'] --> [str(refNum)]
	wireSDInfo = readWireSDInfo(wireE) 
	addRefDict = referenceDictDFromExc['add']
	for nName in referenceDictDFromExc['newRefName']:
		refNameToCopy = addRefDict[nName][0]
		copy = writeARefCopy(refNameToCopy, nName, addRefDict[nName][1], prefix)
		root.insert(int(nName)-1, copy)
		### Change wire des
		if refNameToCopy in wireSDInfo:
			modifyWireDesRef(refNameToCopy, nName, wireE, wireSDInfo[refNameToCopy]['d'], prefix)
	### write to a new xml file
	newXmlFilePath = xmlFolderPath + "/" + xmlFileName + "_new.xml"
	tree.write(newXmlFilePath)
	return newXmlFilePath

def writeARefCopy(oldName, newName, typ, prefix): 
	### creat new referenceSystem node
	newRefEle = Element('ReferenceSystem')
	newNameEle = SubElement(newRefEle, 'Name')
	newNameEle.text = prefix + newName
	newTypeEle = SubElement(newRefEle, 'Type')
	newTypeEle.text = typ
	newDepEle = SubElement(newRefEle, 'Dependon')
	newDepEle.text = prefix + oldName
	### formatting xml text so it prints nicly 
	indent(newRefEle, 1)
	### return the reference(address) of the ref element
	return newRefEle

def modifyWireDesRef(oldDes, newDes, wireElements, wireIndex, prefix): 
	### wireIndex is an array of wire element whose destination needs to be modified
	for index in wireIndex:
		desBond = wireElements[index].findall('Bond')[1].find('Refsys')
		if desBond.text == prefix + oldDes:
			desBond.text = prefix + newDes

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