// MyConnect.h : CMyConnect の宣言

#pragma once
#include "resource.h"       // メイン シンボル



#include "ExcelSimpleComAddin_i.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "DCOM の完全サポートを含んでいない Windows Mobile プラットフォームのような Windows CE プラットフォームでは、単一スレッド COM オブジェクトは正しくサポートされていません。ATL が単一スレッド COM オブジェクトの作成をサポートすること、およびその単一スレッド COM オブジェクトの実装の使用を許可することを強制するには、_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA を定義してください。ご使用の rgs ファイルのスレッド モデルは 'Free' に設定されており、DCOM Windows CE 以外のプラットフォームでサポートされる唯一のスレッド モデルと設定されていました。"
#endif

using namespace ATL;

extern _ATL_FUNC_INFO NewWorkbookInfo;
// CMyConnect

class ATL_NO_VTABLE CMyConnect :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CMyConnect, &CLSID_MyConnect>,
	public ISupportErrorInfo,
	public IMyConnect,
	public IDispatchImpl < _IDTExtensibility2, &__uuidof(_IDTExtensibility2), &LIBID_AddInDesignerObjects, /* wMajor = */ 1 >,
	public IDispEventSimpleImpl<1, CMyConnect, &__uuidof(Excel::AppEvents)>
{
public:
	CMyConnect()
	{
	}

	DECLARE_REGISTRY_RESOURCEID(IDR_MYCONNECT)


	BEGIN_COM_MAP(CMyConnect)
		COM_INTERFACE_ENTRY(IMyConnect)
		COM_INTERFACE_ENTRY(ISupportErrorInfo)
		COM_INTERFACE_ENTRY(_IDTExtensibility2)
	END_COM_MAP()

	// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:




	// _IDTExtensibility2 Methods
public:
	STDMETHOD(OnConnection)(LPDISPATCH Application, ext_ConnectMode ConnectMode, LPDISPATCH AddInInst, SAFEARRAY * * custom)
	{
		CComQIPtr<Excel::_Application> app(Application);

		m_AppPtr = app;
		MessageBox(NULL, L"OnConnection", L"OnConnection", MB_OK);

		HRESULT hr = IDispEventSimpleImpl<1, CMyConnect, &__uuidof(Excel::AppEvents)>::DispEventAdvise(m_AppPtr);
		
		if (FAILED(hr))
			return hr;

		return S_OK;
	}
	STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY * * custom)
	{
		HRESULT hr = IDispEventSimpleImpl<1, CMyConnect, &__uuidof(Excel::AppEvents)>::DispEventUnadvise(m_AppPtr);
		m_AppPtr = NULL;
		return S_OK;
	}
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY * * custom)
	{
		return S_OK;
	}
	STDMETHOD(OnStartupComplete)(SAFEARRAY * * custom)
	{
		return S_OK;
	}
	STDMETHOD(OnBeginShutdown)(SAFEARRAY * * custom)
	{
		return S_OK;
	}

	BEGIN_SINK_MAP(CMyConnect)
		SINK_ENTRY_INFO(/*nID =*/ 1, __uuidof(Excel::AppEvents), /*dispid =*/ 0x0000061d, NewWorkbook, &NewWorkbookInfo)
	END_SINK_MAP()

	void __stdcall NewWorkbook(struct _Workbook * Wb)
	{
		MessageBox(NULL, L"NewWorkbook", L"NewWorkbook", MB_OK);
	}
private:
	CComPtr<Excel::_Application> m_AppPtr;
};

OBJECT_ENTRY_AUTO(__uuidof(MyConnect), CMyConnect)
