import xml.etree.ElementTree as ET
from xml.etree.ElementTree import ElementTree, Element, SubElement
import re
from os.path import exists
from util import splitFileFolderAndName

def readXML(xmlFilePath):
	parseFailure = False
	try:
		tree = ET.parse(xmlFilePath)                                    
	except ET.ParseError: 
		parseFailure  = True

	if not parseFailure:
		root = tree.getroot() 
		referenceE = root.findall('ReferenceSystem')
		wireE = root.findall('Wire')

	if parseFailure or not referenceE or not wireE:
		return None, None, None, None, None

	refName = [] #str
	refNameGap = [] #list of missing ref numbers in str
	typ = []
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
		
	return refName, refNameGap, typ, dependon, wireE

def checkRepeats(refNameList):
	''' Check if there is any repeating reference. Return a list of lists of a name of repeating ref(str) and count(int)'''
	repeat = []
	singles = set(refNameList)
	for s in singles:
		count = refNameList.count(s)
		if count > 1:
			repeat.append([s, count])
	return repeat

def XMLInfo(xmlFilePath, repRef, refName, refGap, wireList):
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
		
def modifier(xmlFilePath, referenceDictDFromExc):
	xmlFolderPath, xmlFileName = splitFileFolderAndName(xmlFilePath)
	### make a ElementTree object and find its root (highest node)   
	try:
		tree = ET.parse(xmlFilePath)                                    
	except ET.ParseError: 
		return None

	root = tree.getroot() 
	### make two lists of all reference elements(objects) and wire elements(objects)
	referenceE = root.findall('ReferenceSystem')
	wireE = root.findall('Wire') 
	numOfRef = len(referenceE)
	numOfWire = len(wireE)
	referenceEDict = {}
	for r in referenceE: ### Modify existing ref's type and dep if necessary
		ref = r.find('Name').text
		typ, dep = referenceDictDFromExc['og'][ref[1:]]
		if r.find('Type').text != typ:
			r.find('Type').text = typ
		if dep:
			if r.find('Dependon') == None:
				newDepEle = Element('Dependon')
				newDepEle.text = 'R' + dep
				r.insert(2, newDepEle)
				indent(newDepEle, 2)
			elif r.find('Dependon').text != 'R'+ dep:
				r.find('Dependon').text = 'R' + dep
		else:
			if r.find('Dependon') != None:
				r.remove(r.find('Dependon'))
		### make reference element dictionary
		referenceEDict[ref] = r

	addRefDict = referenceDictDFromExc['add']
	for nName in referenceDictDFromExc['newRefName']:	# sorted(map(int, addRefDict.keys()))
		copy = writeARefCopy(referenceEDict['R'+addRefDict[nName][0]], addRefDict[nName][0], nName, addRefDict[nName][1])
		root.insert(int(nName)-1, copy)
		### Change wire des
		modifyWireDesRef(addRefDict[nName][0], nName, wireE)
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

def modifyWireDesRef(oldDes, newDes, wireElements):
	for wire in wireElements:
		secBondDes = wire.findall('Bond')[1].find('Refsys')
		if secBondDes.text == 'R' + oldDes:
			secBondDes.text = 'R' + newDes

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










	

