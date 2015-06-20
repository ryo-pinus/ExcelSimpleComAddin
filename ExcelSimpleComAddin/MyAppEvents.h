// MyAppEvents.h : CMyAppEvents の宣言

#pragma once
#include "resource.h"       // メイン シンボル



#include "ExcelSimpleComAddin_i.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "DCOM の完全サポートを含んでいない Windows Mobile プラットフォームのような Windows CE プラットフォームでは、単一スレッド COM オブジェクトは正しくサポートされていません。ATL が単一スレッド COM オブジェクトの作成をサポートすること、およびその単一スレッド COM オブジェクトの実装の使用を許可することを強制するには、_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA を定義してください。ご使用の rgs ファイルのスレッド モデルは 'Free' に設定されており、DCOM Windows CE 以外のプラットフォームでサポートされる唯一のスレッド モデルと設定されていました。"
#endif

using namespace ATL;

extern _ATL_FUNC_INFO NewWorkbookInfo;
// CMyAppEvents

class ATL_NO_VTABLE CMyAppEvents :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CMyAppEvents, &CLSID_MyAppEvents>,
	public IMyAppEvents,
	public IDispEventSimpleImpl<1, CMyAppEvents, &__uuidof(Excel::AppEvents)>
{
public:
	CMyAppEvents()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_MYAPPEVENTS)


BEGIN_COM_MAP(CMyAppEvents)
	COM_INTERFACE_ENTRY(IMyAppEvents)
END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

	// イベントハンドラ
public:
	BEGIN_SINK_MAP(CMyAppEvents)
		SINK_ENTRY_INFO(/*nID =*/ 1, __uuidof(Excel::AppEvents), /*dispid =*/ 0x0000061d, NewWorkbook, &NewWorkbookInfo)
	END_SINK_MAP()

	void __stdcall NewWorkbook(struct _Workbook * Wb);

};

OBJECT_ENTRY_AUTO(__uuidof(MyAppEvents), CMyAppEvents)
