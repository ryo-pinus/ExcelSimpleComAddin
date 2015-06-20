// MyAppEvents.h : CMyAppEvents �̐錾

#pragma once
#include "resource.h"       // ���C�� �V���{��



#include "ExcelSimpleComAddin_i.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "DCOM �̊��S�T�|�[�g���܂�ł��Ȃ� Windows Mobile �v���b�g�t�H�[���̂悤�� Windows CE �v���b�g�t�H�[���ł́A�P��X���b�h COM �I�u�W�F�N�g�͐������T�|�[�g����Ă��܂���BATL ���P��X���b�h COM �I�u�W�F�N�g�̍쐬���T�|�[�g���邱�ƁA����т��̒P��X���b�h COM �I�u�W�F�N�g�̎����̎g�p�������邱�Ƃ���������ɂ́A_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA ���`���Ă��������B���g�p�� rgs �t�@�C���̃X���b�h ���f���� 'Free' �ɐݒ肳��Ă���ADCOM Windows CE �ȊO�̃v���b�g�t�H�[���ŃT�|�[�g�����B��̃X���b�h ���f���Ɛݒ肳��Ă��܂����B"
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

	// �C�x���g�n���h��
public:
	BEGIN_SINK_MAP(CMyAppEvents)
		SINK_ENTRY_INFO(/*nID =*/ 1, __uuidof(Excel::AppEvents), /*dispid =*/ 0x0000061d, NewWorkbook, &NewWorkbookInfo)
	END_SINK_MAP()

	void __stdcall NewWorkbook(struct _Workbook * Wb);

};

OBJECT_ENTRY_AUTO(__uuidof(MyAppEvents), CMyAppEvents)
