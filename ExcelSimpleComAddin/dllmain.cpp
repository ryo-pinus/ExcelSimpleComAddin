// dllmain.cpp : DllMain �̎���

#include "stdafx.h"
#include "resource.h"
#include "ExcelSimpleComAddin_i.h"
#include "dllmain.h"
#include "xdlldata.h"

CExcelSimpleComAddinModule _AtlModule;

// DLL �G���g�� �|�C���g
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
#ifdef _MERGE_PROXYSTUB
	if (!PrxDllMain(hInstance, dwReason, lpReserved))
		return FALSE;
#endif
	hInstance;
	return _AtlModule.DllMain(dwReason, lpReserved); 
}
