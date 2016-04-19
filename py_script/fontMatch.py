#!/usr/bin/python3
# -*- coding: utf-8 -*-

import ctypes
import win32com.client
import os
import codecs
import ctypes
import win32gui
import win32api
import win32con

fontMap = dict();
backupDeleteFont = list();

def deleteFont(fontName):
	if (len(fontName) == 0):
		print("fontName is empty")
		return False

	if (len(fontMap) == 0):
		filePath = os.path.join(os.getcwd(), "config", "fontMap.txt")
		if (not os.path.exists(filePath)):
			from systemFontFiles import getSystemFontAndFontFile
			getSystemFontAndFontFile()

		fontFile = open(filePath, "r")
		for line in fontFile.readlines():
			temp = line.split("||")
			key = temp[0].strip()
			values = [x.strip() for x in temp[1].split("&") if len(x.strip()) > 0]
			fontMap[key] = values

	#print (fontMap)
	print (fontName)
	for (key, values) in fontMap.items():
		for item in values:
			if (item.find(fontName) == 0):
				print (item)
				if(ctypes.windll.gdi32.RemoveFontResourceW(key)):
					print ("delete font: " + key)
					backupDeleteFont.append(key)
					break

	return True

def rellbackFonts():
	#还原
	for item in backupDeleteFont:
		result = ctypes.windll.gdi32.AddFontResourceW(item)
		print ("add font:" + item)
	
	backupDeleteFont.clear()

def saveToFile(fontList):
	print (fontList)
	str = "";
	for item in fontList:
		if (len(item) > 0):
			str += ' __X("'+ item +'"),'

	if (len(str) == 0):
		return

	filePath = os.path.join(os.getcwd(), "config", "result.txt")
	file = codecs.open(filePath, "a", "utf-8")
	file.write("{" + str[:-1] + " },\n")
	file.close()


def genEmfFile(app, pptxFile, fontName, index):
	pres = app.Presentations.open(pptxFile)
	emfFileName = os.path.join(os.getcwd(), "emf", fontName + "_" + str(index) + ".emf")
	pres.slides(1).shapes(1).export(emfFileName, 5)
	pres.close()
	return emfFileName

def getEmfFontName(emfFileName):
	print (emfFileName)

	file = open(os.path.join(os.getcwd(), "dll", "emfFileName.txt"), "w")
	file.write(emfFileName);
	file.close()

	dllPath = os.path.join(os.getcwd(), "dll", "Dllemf.dll")
	hllDll = ctypes.WinDLL(dllPath)
	hllDll.emfFontName()
	
	file = open(os.path.join(os.getcwd(), "dll", "emfFontName.txt"), "r")
	fontName = file.readline().strip();
	file.close();
	print (fontName)
	return fontName

def getEndFontList():
	file = codecs.open(os.path.join(os.getcwd(), "config", "endFontList.txt"), "r", "utf-8")
	fonts = []
	for line in file.readlines():
		fontName = line.strip()
		if len(fontName):
			fonts.append(fontName)

	file.close()
	return fonts

def maching():
	endFonts = getEndFontList()
	print(endFonts)
	filePath = os.path.join(os.getcwd(), "pptxfiles")
	listDir = os.listdir(filePath)
	if (len(listDir) == 0):
		from generPptxFiles import generFiles
		generFiles(os.path.join(os.getcwd(), "config", "testfonts.txt"))
		listDir = os.listdir(filePath)
	print (listDir)
	for file in listDir:
		print (file)
		index = 1
		fileName= os.path.basename(file).split(".")[0]
		absFilePath = os.path.join(filePath, file)
		print (absFilePath)
		app = win32com.client.Dispatch('Powerpoint.Application')
		emfFileName = genEmfFile(app, absFilePath, fileName, index)
		index = index + 1
		app.Quit()
		fontList = []
		fontName = getEmfFontName(emfFileName)
		while ((fontName not in fontList)):
			fontList.append(fontName)
			if (deleteFont(fontName) == False):
				break

			#判断是否可以停止了
			canStop = [ x for x in endFonts if x in fontList ]
			print (canStop)
			if (len(canStop)):
				break

			app = win32com.client.Dispatch('Powerpoint.Application')
			emfFileName = genEmfFile(app, absFilePath, fileName, index)
			index = index + 1
			app.Quit()
			fontName = getEmfFontName(emfFileName)

		saveToFile(fontList)
		rellbackFonts()


if __name__ == '__main__':
	maching();