import xml.etree.ElementTree as ET
from xml.etree.ElementTree import ElementTree, Element, SubElement
import re
from os.path import exists
from util import splitFileFolderAndName

def readXML(xmlFilePath):
	"""Read a XML data file and collect reference and wire information.

		Parameters
		----------
		xmlFilePath: string

		Returns
		-------
		referenceInfo: dictionary
			referenceInfo consists of lists of reference names, types, dependencies, gaps, and repeats.
			referenceInfo --> {'name': ['refNum'], 'type': ['type'], 'dependon': ['refNum'], 'gap': ['refNum'],'repeats': [['refname', number]]}

		wireSDCount: dictionary
			keys: All reference names. 
			values: Dictionaries of wire source and destination count. 
	"""

	parseFailure = False
	try:
		# Parse xml content
		tree = ET.parse(xmlFilePath)                                    
	except ET.ParseError: 
		# True if it fails parsing
		parseFailure  = True
	# Get root of the xml data tree and all reference and wire elements if xml is parsed successfully
	if not parseFailure:
		root = tree.getroot() 
		referenceE = root.findall('ReferenceSystem')
		wireE = root.findall('Wire')
	# if there is an error in parsing or getting elements, return empty dictionaries
	if parseFailure or not referenceE or not wireE:
		return {}, {}

	refName = [] 				# A list of reference numbers in string
	refNameGap = [] 			# A list of missing reference numbers in string
	typ = [] 					# A list of refernece types in string
	dependon = [] 				# A list of reference number that references depend on in string
	numOfRef = len(referenceE)	# The number of reference elements

	# Obtain wire source and destination information and total count 
	wireSDCount = readWireSDInfo(wireE)
	
	# Obtain ref, type, dep, and gaps
	for i in range(0, numOfRef):
		# Get reference numbers in string without the prefix('R')
		numberS = re.findall('\d+', referenceE[i].find('Name').text)[0]
		# Get dependent values in string
		depS = referenceE[i].find('Dependon')
		# If dependent values exist, get dependent numbers in string without the prefix('R')
		if depS != None:
			depS = re.findall('\d+', depS.text)[0]
		# Append reference name, type, and dependency to lists
		refName.append(numberS)
		typ.append(referenceE[i].find('Type').text)
		dependon.append(depS)
		# Obtain gaps if difference between current reference number and previous is greater than 1
		if i > 0:
			currNum = int(numberS)
			if currNum - prevNum > 1:
				for missing in range(prevNum + 1, currNum):
					# Append refernce name in string to list
					refNameGap.append(str(missing))
			prevNum = currNum
		else:
			prevNum = int(numberS)
		# Add on references has no wires attached to
		if not numberS in wireSDCount:
			wireSDCount[numberS] = {'s': [], 'd': []}

	# Compress references information into a dictionary
	referenceInfo = {'name': refName,'type': typ, 'dependon': dependon, 'gap': refNameGap, 'repeats': checkRepeats(refName)}
	
	return referenceInfo, wireSDCount

def readWireSDInfo(wireElements):
	"""Obtain wire source and destination counts.

		Parameters
		----------
		wireElements: list
			A list of wire elements from a parsed XML file

		Returns
		-------
	 	wireSDInfo: dictionary
	 		A dictionary contains information such as total number of wires and dictionaries for lists of wire indices for sources and destinations
	 		that references are treated as. It does not include references that has no wire attached to. 
	 		wireSDInfo --> {totalWireCount: number in int, 'reference number in str': {s:[wire index], d:[wire index]}}
	"""
	# Get total number of wires
	wireSDInfo = {'total': len(wireElements)}
	for wireIndex in range(0, len(wireElements)):
		# Get reference number in string without prefix('R') for source
		source = re.findall('\d+', wireElements[wireIndex].findall('Bond')[0].find('Refsys').text)[0] 
		# Get reference number in string without prefix('R') for desination
		destination = re.findall('\d+', wireElements[wireIndex].findall('Bond')[1].find('Refsys').text)[0]
		# Add reference name as source into the dictionary with wire index
		if source in wireSDInfo:
			wireSDInfo[source]['s'].append(wireIndex)
		else:
			wireSDInfo[source] = {'s': [wireIndex], 'd': []}
		# Add reference name as destination into the dictionary with wire index
		if destination in wireSDInfo:
			wireSDInfo[destination]['d'].append(wireIndex)
		else:
			wireSDInfo[destination] = {'s': [], 'd': [wireIndex]}
	return wireSDInfo

def checkRepeats(refNameList):
	"""Check for any repeating references.
		
		Parameters
		----------
		refNameList: list
			A list of reference name in the XML file.

		Returns
		-------
		repeat: list
			1. A list of lists of reference names in str and number of times they repeat. 
			2. A empty list when there is no repeats
			repeat --> [['reference number', count]]
	"""

	repeat = []
	# Get a set of reference name which has no duplicates
	singles = set(refNameList)
	# If both has the same number of elements, there is no repeats
	if len(refNameList) == len(singles):
		return repeat 
	# Append to list when the count of reference name is more than one
	for s in singles:
		count = refNameList.count(s)
		if count > 1:
			repeat.append([s, count])
	return repeat

def XMLInfo(xmlFilePath, repRef, refName, refGap, wireCount):
	"""Creat a message of all information of the XML data. It's typically used when there is an error in the XML file.

		Parameters
		----------
		xmlFilePath: string
			XML file path.
		repRef: list
			A list of lists of repeating reference names and their count.
		refName: list
			A list of reference name in string.
		refGap: list
			A list of missing refernece name(gap) in string.
		wireCount: int
			The number of wires.

		Returns
		-------
		info: string
			Information of the XML file.
	"""

	prefix = 'R'	

	if exists(xmlFilePath):
		info = ""
		# Write file path		
		info = info + "#Input XML File: " + xmlFilePath + '\n\n'
		# Write repeating ref name if there is any
		info = info + "#Repeating Reference:\n"
		if repRef:
			for r in repRef:
				info = info + "There are " + str(r[1]) + " R" + r[0] + '\n'
		else:
			info  = info + "None\n"
		# Write first and last ref name
		info = info + "\n#First Reference: R" + refName[0] + '\n'
		info = info + "#Last Reference:  R" + refName[-1] + '\n'
		# Write refernce gaps
		info = info + "\n#Range of Gaps (included):\n"
		if refGap:
			for g in refGap:
				if len(g) == 1:
					info = info + prefix + str(g[0]) + '\n'
				else:
					info = info + prefix + str(g[0]) + ' - ' + prefix + str(g[1]) + '\n'
		else:
			info = info + "None\n"
		# Write wire count
		info = info + "\n#Number of Wires: " + str(wireCount) + "\n"
	else:
		info = "File does not exist!"

	return info
		
def modifier(xmlFilePath, referenceDictDFromExc):
	"""Create a new XML file with modified information.
		
		Parameters
		----------
		xmlFilePath: string
			XML file path.
		referenceDictDFromExc: dictionary
			A collection of data from an Excel spreadsheet.

		Returns
		-------
		message: string
			An error message.

		or

		newXmlFilePath: string
			A new XML file path.
	"""

	# Split a XML file path into folder path and file name without the extension
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
	"""Create a referece element (no points).

		Parameters
		----------
		oldName: string
			The reference name to copy
		newNmea: string
			Name of the new reference.
		typ: string
			Type of new reference.
		prefix: string
			String before number in reference name.

		Returns
		-------
		newRefEle: Element
			New reference element.
	"""

	# Creat a new referenceSystem node
	newRefEle = Element('ReferenceSystem')
	# Creat a sub-element for name in reference
	newNameEle = SubElement(newRefEle, 'Name')
	newNameEle.text = prefix + newName
	# Creat a sub-element for type in reference
	newTypeEle = SubElement(newRefEle, 'Type')
	newTypeEle.text = typ
	# Creat a sub-element for dependon in reference
	newDepEle = SubElement(newRefEle, 'Dependon')
	newDepEle.text = prefix + oldName
	# Formatting xml text so it prints nicly 
	indent(newRefEle, 1)
	# Return the reference(address) of the new reference element
	return newRefEle

def modifyWireDesRef(oldDes, newDes, wireElements, wireIndex, prefix): 
	"""Modify wire destinations.

		Parameters
		----------
		oldDes: string
			The original destination (reference number in string). 
		newDes : string
			The new destination (reference number in string). 
		wireElements: Element
			Wire elements in the XML file
		wireIndex: list
			A list of wire indices of wire elements whose destinations need to be modified.
		prefix: string
			String before number in reference name.
	"""

	for index in wireIndex:
		# Find destination element in a wire element 
		desBond = wireElements[index].findall('Bond')[1].find('Refsys')
		# Change the name from oldDes to newDes
		if desBond.text == prefix + oldDes:
			desBond.text = prefix + newDes

def indent(elem, level=0):
	"""In-place prettyprint formatter found online --> http://effbot.org/zone/element-lib.htm #prettyprint."""

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