// MyConnect.cpp : CMyConnect ‚ÌŽÀ‘•

#include "stdafx.h"
#include "MyConnect.h"
#include "MyAppEvents.h"


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


STDMETHODIMP CMyConnect::OnConnection(LPDISPATCH Application, ext_ConnectMode ConnectMode, LPDISPATCH AddInInst, SAFEARRAY * * custom)
{
	CComQIPtr<Excel::_Application> app(Application);

	m_AppPtr = app;
	MessageBox(NULL, L"OnConnection", L"OnConnection", MB_OK);

	
	IMyAppEvents *pAppEvents = NULL;
	CMyAppEvents::CreateInstance<>(&pAppEvents);
	m_appEvents = pAppEvents;

	HRESULT hr = ((CMyAppEvents*)m_appEvents)->DispEventAdvise(m_AppPtr);
	if (FAILED(hr))
		return hr;

	return S_OK;
}
STDMETHODIMP CMyConnect::OnDisconnection(ext_DisconnectMode RemoveMode, SAFEARRAY * * custom)
{
	HRESULT hr = ((CMyAppEvents*)m_appEvents)->DispEventUnadvise(m_AppPtr);
	m_appEvents->Release();
	m_appEvents = NULL;
	m_AppPtr = NULL;
	return S_OK;
}
STDMETHODIMP CMyConnect::OnAddInsUpdate(SAFEARRAY * * custom)
{
	return S_OK;
}
STDMETHODIMP CMyConnect::OnStartupComplete(SAFEARRAY * * custom)
{
	return S_OK;
}
STDMETHODIMP CMyConnect::OnBeginShutdown(SAFEARRAY * * custom)
{
	return S_OK;
}
