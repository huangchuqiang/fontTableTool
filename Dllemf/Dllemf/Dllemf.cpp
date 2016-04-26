// Dllemf.cpp : 定义 DLL 应用程序的导出函数。
//

#include "stdafx.h"
#include "Dllemf.h"

#include "TTF.h"
#include "TTC.h"
#include <vector>
#include <string>
#include<fstream>
#include <iostream>
#include <cctype>
#include <algorithm>
#include <functional>

#include <string>
#include<atlstr.h>
#include "gel/stxutif.h"

#include <windows.h>
#pragma comment(lib, "kernel32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#define OUT

using namespace std;
using namespace gel;

static bool scanEmfHeader(const char* bits, const int len, OUT int& offset)
{
	offset = 0;

	if (offset + sizeof(EMR) > len)
		return false;
	const EMR* pEmr = (const EMR*)bits;

	if (EMR_HEADER != pEmr->iType)
		return false;

	if (offset + pEmr->nSize > len)
		return false;

	const ENHMETAHEADER* pEmfHeader = (const ENHMETAHEADER*)bits;
	offset += pEmr->nSize;

	if (pEmfHeader->dSignature != ENHMETA_SIGNATURE ||
		pEmfHeader->nRecords < 2 ||
		0 != (pEmfHeader->nSize % 4))	// 4 bytes align
	{
		return false;
	}

	return true;
}

bool readFile(IN const char* filepath, OUT std::vector<char>& bits)
{
	ifstream in(filepath, ios::in | ios::binary | ios::ate);
	long size = in.tellg();
	if (size <= 0)
		return false;

	in.seekg(0, ios::beg);
	bits.resize(size);
	in.read(bits.data(), size);
	in.close();

	return true;
}

static std::string scanFont(char* bits, const int len)
{
	int offset = 0;
	if (!scanEmfHeader(bits, len, offset))
		return std::string();

	const char* pData = bits + offset;
	const char* end = bits + len;
	EMR* pEmr = nullptr;

	for (; pData < end;)
	{
		if (pData + sizeof(EMR) > end)
			break;

		pEmr = (EMR*)pData;

		if (!pEmr->nSize ||
			pData + pEmr->nSize > end)
			break;

		if (pEmr->iType > EMR_HEADER &&
			pEmr->iType < EMR_MAX &&
			sizeof(EMR) <= pEmr->nSize &&
			(pEmr->nSize % 4 == 0))
		{
			if (EMR_EOF == pEmr->iType)
				break;
			else if (EMR_EXTCREATEFONTINDIRECTW == pEmr->iType)
			{
				EMREXTCREATEFONTINDIRECTW* pEmr = (EMREXTCREATEFONTINDIRECTW*)pData;
				std::wstring strText(pEmr->elfw.elfLogFont.lfFaceName);
				if (!strText.empty())
					return ws2s(strText);
			}
		}

		pData += pEmr->nSize;
	}

	return std::string();
}

std::string getEmfFont(const char* src)
{
	vector<char> srcBits;
	readFile(src, srcBits);
	if (srcBits.empty())
		return std::string();

	//std::cout << "scan";
	return scanFont(srcBits.data(), srcBits.size());
}

DLLEMF_API bool emfFontName()
{
	ifstream fin(dllPath + "dll\\emfFileName.txt");
	char fileName[100] = { 0 };
	fin >> fileName;
	//std::cout << 1;

	std::string strFontName(getEmfFont(fileName));
	//std::cout << 1;
	if (strFontName.empty())
		return false;

	//std::cout << 1;
	ofstream fout(dllPath + "dll\\emfFontName.txt");
	fout << strFontName.c_str();
	fout.close();
	return true;
}

//-------------------------------------------------------------------------------------------
void getTtfFileName(vector<wstring>& vecFontsFiles)
{
	std::wifstream fsi;
	fsi.imbue(stdx::utf8_locale);
	fsi.open(dllPath + "config\\fontTables.txt");
	cout << "begin open fontTables.txt" << std::endl;
	if (!fsi.good())
	{
		std::cout << "open fontTables.txt error" << std::endl;
		return;
	}
	wchar_t c; 
	while (!fsi.eof())
	{
		std::wstring s;
		fsi >> s;
		if (!s.empty())
			vecFontsFiles.push_back(__T("c:\\windows\\fonts\\") + s);
	}

	fsi.close();
	cout << "open fontTables ok" << std::endl;
}

void saveFonts(vector<wstring>& vecFontsName, vector<wstring>& vecFiles)
{
	std::wofstream fso;
	fso.imbue(stdx::utf8_locale);
	fso.open(dllPath + "config\\fontMap.txt", fso.out | fso.trunc);
	cout << "begin save fontMap" << std::endl;
	if (!fso.good())
	{
		cout << "save fontMap error" << std::endl;
		return;
	}

	std::size_t len = vecFiles.size();
	for (int i = 0; i < len; ++i)
		fso << vecFiles[i].c_str() << " || " << vecFontsName[i].c_str() << std::endl;

	fso.close();
	cout << "save fontMap ok" << std::endl;
}

DLLEMF_API bool getFontNameFromTTF()
{
	vector<wstring> ttfFiles;
	getTtfFileName(ttfFiles);
	vector<wstring> ttfFonts;

	for (auto it = ttfFiles.begin(); it != ttfFiles.end(); ++it)
	{
		std::wstring file(*it);
		if (file.find(__T(".TTF")) != string::npos || file.find(__T(".ttf")) != std::string::npos)
		{
			TTF ttf;
			if (ttf.Parse(file))
				ttfFonts.push_back(ttf.GetFontName());
			else
				ttfFonts.push_back(file);
		}
		else if (file.find(__T(".TTC")) != string::npos || file.find(__T(".ttc")) != string::npos)
		{
			TTC ttc;
			if (ttc.Parse(file))
			{
				wstring fontName;
				for (int i = 0; i < ttc.Size(); ++i)
					fontName += ttc.GetFontName(i) + __T(" & ");

				ttfFonts.push_back(fontName);
			}
		}
		else
		{
			ttfFonts.push_back(file);
		}
	}

	saveFonts(ttfFonts, ttfFiles);
	return true;
}