// ExTVw.cpp: Implementation of DLL exports.

#include "stdafx.h"
#include "res/resource.h"
#ifdef UNICODE
	#include "ExTVwU.h"
#else
	#include "ExTVwA.h"
#endif


class ExTVwModule :
    public CAtlDllModuleT<ExTVwModule>
{

public:
  #ifdef UNICODE
		DECLARE_LIBID(LIBID_ExTVwLibU)
		DECLARE_REGISTRY_APPID_RESOURCEID(IDR_EXTVW, "{26AB9AEF-6314-4D4D-8F4D-BB4F5F6C167C}")
	#else
		DECLARE_LIBID(LIBID_ExTVwLibA)
		DECLARE_REGISTRY_APPID_RESOURCEID(IDR_EXTVW, "{A721B8EE-E9EC-4359-96BE-B04D8F410344}")
	#endif
};

ExTVwModule _AtlModule;


#ifdef _MANAGED
	#pragma managed(push, off)
#endif

// DLL entry point
extern "C" BOOL WINAPI DllMain(HINSTANCE /*hInstance*/, DWORD dwReason, LPVOID lpReserved)
{
	#ifdef _DEBUG
		// enable CRT memory leak detection & report
		_CrtSetDbgFlag(_CRTDBG_ALLOC_MEM_DF     // enable debug heap allocs & block type IDs (ie _CLIENT_BLOCK)
		    //| _CRTDBG_CHECK_CRT_DF              // check CRT allocations too
		    //| _CRTDBG_DELAY_FREE_MEM_DF       // keep freed blocks in list as _FREE_BLOCK type
		    | _CRTDBG_LEAK_CHECK_DF             // do leak report at exit (_CrtDumpMemoryLeaks)

		    // pick only one of these heap check frequencies
		    //| _CRTDBG_CHECK_ALWAYS_DF         // check heap on every alloc/free
		    //| _CRTDBG_CHECK_EVERY_16_DF
		    //| _CRTDBG_CHECK_EVERY_128_DF
		    //| _CRTDBG_CHECK_EVERY_1024_DF
		    | _CRTDBG_CHECK_DEFAULT_DF          // by default, no heap checks
		    );
		//_CrtSetBreakAlloc(209);               // break debugger on numbered allocation
		// get ID number from leak detector report of previous run
	#endif

	return _AtlModule.DllMain(dwReason, lpReserved); 
}

#ifdef _MANAGED
	#pragma managed(pop)
#endif

STDAPI DllCanUnloadNow(void)
{
	return _AtlModule.DllCanUnloadNow();
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	return _AtlModule.DllGetClassObject(rclsid, riid, ppv);
}

STDAPI DllRegisterServer(void)
{
	return _AtlModule.DllRegisterServer();
}

STDAPI DllUnregisterServer(void)
{
	return _AtlModule.DllUnregisterServer();
}