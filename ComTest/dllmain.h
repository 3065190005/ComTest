// dllmain.h: 模块类的声明。

class CComTestModule : public ATL::CAtlDllModuleT< CComTestModule >
{
public :
	DECLARE_LIBID(LIBID_ComTestLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_COMTEST, "{18f82e34-9b50-4e39-94f5-9dc937cf6d57}")
};

extern class CComTestModule _AtlModule;
