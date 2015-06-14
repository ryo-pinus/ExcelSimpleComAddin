// MyConnect.cpp : CMyConnect ‚ÌŽÀ‘•

#include "stdafx.h"
#include "MyConnect.h"


// CMyConnect

STDMETHODIMP CMyConnect::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* const arr[] = 
	{
		&IID_IMyConnect
	};

	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}
