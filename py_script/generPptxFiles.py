#!/usr/bin/python3
# -*- coding: utf-8 -*-

import win32com.client
import os
import codecs

def generFiles(fileName):
	"""
		传进来一个文件名，从里面读取要测试的字体列表。给这些字体生成一个相应的pptx文件
	"""

	if (not os.path.exists(fileName)):
		print ("cannot get test fonts")
		return

	file = codecs.open(fileName, "r", "utf-8")
	fonts = []
	for line in file.readlines():
		fonts.append(line.strip())

	file.close()
	print(fonts)

	app = win32com.client.Dispatch('Powerpoint.Application')

	for item in fonts:
		pres = app.Presentations.Add()
		slide = pres.slides.Add(1, 1) #ppLayoutTitle
		slide.shapes(1).TextFrame2.TextRange.Text = "Aa"
		#设置字体
		slide.shapes(1).TextFrame2.TextRange.Font.Name = item
		slide.shapes(1).TextFrame2.TextRange.Font.NameFarEast = item
		slide.shapes(1).TextFrame2.TextRange.Font.NameAscii = item
		pres.saveas(os.path.join(os.getcwd(), "pptxfiles", item + ".pptx"))
		pres.close()

	app.Quit()

if __name__ == '__main__':
	generFiles(os.path.join(os.getcwd(), "config", "testfonts.txt"))

