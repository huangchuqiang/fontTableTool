#!/usr/bin/python3
# -*- coding: utf-8 -*-

import os
import types
import codecs 
import ctypes
import win32api

def getFontName():
	#取字体ttf文件里面的字体名
	dllPath = os.path.join(os.getcwd(), "dll", "Dllemf.dll")
	hllDll = ctypes.WinDLL(dllPath)
	print (hllDll)
	hllDll.getFontNameFromTTF()

def getSystemFontAndFontFile():
	"""
	从系统注册表中取到安装的字体及对应的ttf文件
	得到的结果写到文件里面
	"""

	import winreg
	key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 'software\Microsoft\Windows NT\CurrentVersion\Fonts', 0, winreg.KEY_READ)
	result = []
	index = 0
	while True:
		try:
			valueList = winreg.EnumValue(key, index)
			result.append(valueList)
			index += 1
		except OSError:
			break
	winreg.CloseKey(key)
	print(result)

	filePath = os.path.join(os.getcwd(),"config", "fontTables.txt")
	fontFile = codecs.open(filePath, "w", "utf-8")
	for items in result:
		#print (items)
		fontFile.write(items[1] + " ")

	fontFile.close()

	getFontName()



if __name__ == '__main__':
	getSystemFontAndFontFile()