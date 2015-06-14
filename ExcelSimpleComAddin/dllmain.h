// dllmain.h : モジュール クラスの宣言

class CExcelSimpleComAddinModule : public ATL::CAtlDllModuleT< CExcelSimpleComAddinModule >
{
public :
	DECLARE_LIBID(LIBID_ExcelSimpleComAddinLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_EXCELSIMPLECOMADDIN, "{BD9010A1-6D29-4398-9F09-9752D0948B31}")
};

extern class CExcelSimpleComAddinModule _AtlModule;
