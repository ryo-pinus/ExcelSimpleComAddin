// MyConnect.h : CMyConnect �̐錾

#pragma once
#include "resource.h"       // ���C�� �V���{��



#include "ExcelSimpleComAddin_i.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "DCOM �̊��S�T�|�[�g���܂�ł��Ȃ� Windows Mobile �v���b�g�t�H�[���̂悤�� Windows CE �v���b�g�t�H�[���ł́A�P��X���b�h COM �I�u�W�F�N�g�͐������T�|�[�g����Ă��܂���BATL ���P��X���b�h COM �I�u�W�F�N�g�̍쐬���T�|�[�g���邱�ƁA����т��̒P��X���b�h COM �I�u�W�F�N�g�̎����̎g�p�������邱�Ƃ���������ɂ́A_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA ���`���Ă��������B���g�p�� rgs �t�@�C���̃X���b�h ���f���� 'Free' �ɐݒ肳��Ă���ADCOM Windows CE �ȊO�̃v���b�g�t�H�[���ŃT�|�[�g�����B��̃X���b�h ���f���Ɛݒ肳��Ă��܂����B"
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
