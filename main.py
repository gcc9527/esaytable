# coding: utf-8


from gc import set_debug
from typing import Text
import xlrd
from xlrd import sheet
import os
import string

s=""
ss = ""

def getDictData(dictData):
	key = ""
	val = ""
	for k in dictData:
		key = k
		val = dictData[k]
		break
	return key, val

def parseTab(arr, txt, jsonTxt):
	arrLen = len(arr)
	if arrLen % 2 != 0:
		raise Exception("arrLen % 2 != 0") 

	arrTxt = ""
	jsonArrTxt = ""
	maxLen = arrLen - 2
	for i in range(0, arrLen, 2):
		addSig = True
		inerTxt = "{"
		jsonInerTxt = "{"
			
		dictKey1, dictVal1 = getDictData(arr[i])
		dictKey2, dictVal2 = getDictData(arr[i + 1])
		inerTxt += dictKey1 + "=" + str(dictVal1) + "," + dictKey2 + "=" + str(dictVal2) + "}"
		jsonInerTxt += '"' + dictKey1 + '"' + ":" + str(dictVal1) + "," + '"' + dictKey2 + '"' + ":" + str(dictVal2) + "}"

		if i == maxLen:
			addSig = False
		if addSig:
			inerTxt += ","	
			jsonInerTxt += ","	

		arrTxt += inerTxt
		jsonArrTxt += jsonInerTxt

	txt += arrTxt + "}"
	jsonTxt += jsonArrTxt + "]"

	tab = ""
	arr.clear()	
	ok = True		
	return txt, tab, jsonTxt

def parseData(row, col, sheetDatas, names, types, txt, val, tab, arr, jsonTxt):
	if val == "":
		return txt, tab, arr, jsonTxt

	strCol = str(col)	
	fieldType = types.get(strCol, "")
	fidldName = names.get(strCol, "")
	if fieldType == "" or fidldName == "":
		return txt, tab, arr, jsonTxt

	if col >= 2 and tab == "":
		txt += ","
		jsonTxt += ","

	idx = str.find(fidldName, ".")
	if idx != -1:
		pre = fidldName[:idx]
		erp = fidldName[idx+1::]

		if tab and pre != tab:
			txt, tab, jsonTxt = parseTab(arr, txt, jsonTxt)
			txt += ","
			jsonTxt += ","

		if tab == "":
			tab = pre
			txt += pre + "=" + "{"
			jsonTxt += '"' + pre + '"' + ":" + "["
		
		arr.append({erp:int(val)})
	else:
		ok = False
		if tab != "" and arr:
				txt, tab, jsonTxt = parseTab(arr, txt, jsonTxt)
				ok = True

		vals = ""
		jsonJson = ""
		if fieldType == "int":
			vals = int(val)
		elif fieldType == "string":
			vals += '"' + str(val) + '"'
		elif fieldType == "json":
			listVar = list(val)
			if len(val) > 3:
				if listVar[len(listVar) - 2] == ",":
					listVar.pop(len(listVar) - 2)
			vals = "{"
			jsonJson = "["
			listVarLen = len(listVar)
			if listVarLen == 2:
				vals = "{}"
				jsonJson = "[]"
			else:
				listVar.pop(0)
				listVarLen = len(listVar)
				listVar.pop(listVarLen - 1)
				listVar = "".join(listVar)
				listVar = listVar.split(",")
				listVarLen = len(listVar)
				for i in range(0, listVarLen):
					vals += listVar[i]
					jsonJson += listVar[i]
					if i + 1 < listVarLen:
						vals += ","
						jsonJson += ","
				vals += "}"
				jsonJson += "]"
		if ok:
			txt += ","	
			jsonTxt += ","
		txt += fidldName + "=" + str(vals)
		if jsonJson != "":
			vals = jsonJson
		jsonTxt += '"' + fidldName + '"' + ":" + str(vals)
	return txt, tab, arr, jsonTxt

def parseSheet(sheets, sheet):
	global s
	global ss

	sheetDatas = sheets.sheet_by_name(sheet)
	names = {}
	types = {}

	for row in range(sheetDatas.nrows):
		if row == 0 or sheetDatas.cell_value(row, 0) == "":
			continue
		elif row == 1:
			for col in range(sheetDatas.ncols):
				colVal = sheetDatas.cell_value(row, col)
				if colVal != "":
					names[str(col)] = colVal
		elif row == 2:
			for col in range(sheetDatas.ncols):
				colVal = sheetDatas.cell_value(row, col)
				if colVal != "":
					types[str(col)] = colVal
		else:
			if row >= 4:
				s += ","
				s += "\r\t"
				ss += ","
				ss += "\r\t"

			txt = ""		
			arr = []	
			tab = ""
			jsonTxt = ""
			
			for col in range(sheetDatas.ncols):
				val = sheetDatas.cell_value(row, col)
				if val == "":
					continue
				if col == 0:
					val = int(val)
					txt += "[" + str(val) + "]={"
					jsonTxt +=  '"' + str(val) + '"' + ":{"	
					keyName = names.get("0", "")
					txt += keyName + "=" + str(val) + ","
					jsonTxt += '"' + keyName + '"' + ":" + str(val) + ","				
				else:
					txt, tab, arr, jsonTxt = parseData(row, col, sheetDatas, names, types, txt, val, tab, arr, jsonTxt)

			if tab != "" and arr:
				arrLen = len(arr)
				maxLen = arrLen - 2
				if arrLen % 2 != 0:
					raise Exception("arrLen % 2 != 0") 

				arrTxt = ""
				jsonArrTxt = ""

				for i in range(0, arrLen, 2):
					addSig = True
					inerTxt = "{"
					jsonInerTxt = "{"
					
					dictKey1, dictVal1 = getDictData(arr[i])
					dictKey2, dictVal2 = getDictData(arr[i + 1])
					inerTxt += dictKey1 + "=" + str(dictVal1) + "," + dictKey2 + "=" + str(dictVal2) + "}"
					jsonInerTxt += '"' + dictKey1 + '"' + ":" + str(dictVal1) + "," + '"' + dictKey2 + '"' + ":" + str(dictVal2) + "}"

					if i == maxLen:
						addSig = False
					if addSig:
						inerTxt += ","
						jsonInerTxt += ","	


					arrTxt += inerTxt
					jsonArrTxt += jsonInerTxt
				txt += arrTxt + "}"
				jsonTxt += jsonArrTxt + "]"
				
				tab = ""
				arr.clear()	

			txt += "}"
			jsonTxt += "}"
			s += txt
			ss += jsonTxt

	s += "\n}"		
	ss += "\n}"			

			

def writeFile(fileName):
	sheets = xlrd.open_workbook(fileName)

	for sheet in sheets.sheet_names():
		idx = sheet.find("@")
		if idx == -1:
			continue

		global s
		global ss
		s +="{\n"+ "\t"
		ss +="{\n"+ "\t"

		parseSheet(sheets, sheet)

		name = sheet[idx + 1::]
		expath = r"../Game/Bin/Lua/GameSer/Config/"
		expath += name + ".lua"
		fp = open(expath, "w+", encoding="utf-8")
		ret ="return "
		ret += s
		fp.write(ret)
		fp.close()
		expath = r"../Game/Config/"
		expath += name + ".json"
		fp = open(expath, "w+", encoding="utf-8")
		ret = ""
		ret += ss
		fp.write(ret)
		fp.close()		
		s = ""
		ss = ""
			


def main():
	for _, dirs, files in os.walk("G:\OneDrive\pp"):
		for f in files:
			idx = str.find(f, ".")
			if idx != -1:
				name = f[idx + 1::]
				if name == "xlsx" or name == "xls":
					if str.find(f, "~") == -1 and str.find(f, "@") != -1:
						writeFile(f)
def f1():
	a="{1,2,3,}"
	al=list(a)
	print(al)
	al.pop(len(a)-2)
	a="".join(al)
	print(a)
#f1()		
main()

print("!!!!!!!!!!!!!!!!!!!!!!!!!导表完毕!!!!!!!!!!!!!!!!!!")
print()
print()
print()
print()
os.system("pause")

