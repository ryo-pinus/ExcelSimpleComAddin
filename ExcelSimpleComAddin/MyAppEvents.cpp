// MyAppEvents.cpp : CMyAppEvents ‚ÌŽÀ‘•

#include "stdafx.h"
#include "MyAppEvents.h"


// CMyAppEvents
_ATL_FUNC_INFO NewWorkbookInfo = { CC_STDCALL, VT_EMPTY, 1, { VT_PTR } };

void __stdcall CMyAppEvents::NewWorkbook(struct _Workbook * Wb)
{
	MessageBox(NULL, L"NewWorkbook", L"NewWorkbook", MB_OK);
}

