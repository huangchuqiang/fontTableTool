// dllmain.cpp : 定义 DLL 应用程序的入口点。
#include "stdafx.h"

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
					 )
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
	{
		char FileName[MAX_PATH + 1] = { 0 };
		GetModuleFileNameA(hModule, FileName, MAX_PATH);
		std::string strFile(FileName);
		int index = strFile.rfind("dll\\");
		dllPath = strFile.substr(0, index);
		break;
	}
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
	case DLL_PROCESS_DETACH:
		break;
	}
	return TRUE;
}

