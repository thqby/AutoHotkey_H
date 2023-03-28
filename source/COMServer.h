// COMServer.h: Deklaration von AuotHotkey2
#pragma once
#ifdef _USRDLL
#include <sdkddkver.h>

#define _ATL_APARTMENT_THREADED

#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// Einige CString-Konstruktoren sind explizit.


#define ATL_NO_ASSERT_ON_DESTROY_NONEXISTENT_WINDOW

#include <atlbase.h>
#include <atlcom.h>
#include <atlctl.h>

#include "resources/resource.h"

#include "ComServer_h.h"


#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Singlethread-COM-Objekte werden auf der Windows CE-Plattform nicht vollständig unterstützt. Windows Mobile-Plattformen bieten beispielsweise keine vollständige DCOM-Unterstützung. Definieren Sie _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA, um ATL zu zwingen, die Erstellung von Singlethread-COM-Objekten zu unterstützen und die Verwendung eigener Singlethread-COM-Objektimplementierungen zu erlauben. Das Threadmodell in der RGS-Datei wurde auf 'Free' festgelegt, da dies das einzige Threadmodell ist, das auf Windows CE-Plattformen ohne DCOM unterstützt wird."
#endif

using namespace ATL;

class CAutoHotkeyModule : public ATL::CAtlDllModuleT< CAutoHotkeyModule >
{
public:
	DECLARE_LIBID(LIBID_AutoHotkey)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_AUTOHOTKEY, "{d346977b-cc9d-4428-9cd2-1b41652586fc}")

};
extern class CAutoHotkeyModule _AtlModule;

// AuotHotkey2

class AuotHotkey2 :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<AuotHotkey2, &CLSID_AuotHotkey2>,
	public ISupportErrorInfo,
	public IObjectWithSiteImpl<AuotHotkey2>,
	public IDispatchImpl<ICOMServer, &IID_ICOMServer, &LIBID_AutoHotkey, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	AuotHotkey2();
	//~AuotHotkey2();

	DECLARE_REGISTRY_RESOURCEID(IDR_COMSERVER)


	BEGIN_COM_MAP(AuotHotkey2)
		COM_INTERFACE_ENTRY(ICOMServer)
		COM_INTERFACE_ENTRY(IDispatch)
		COM_INTERFACE_ENTRY(ISupportErrorInfo)
		COM_INTERFACE_ENTRY(IObjectWithSite)
		//COM_INTERFACE_ENTRY_AGGREGATE(IID_IMarshal, m_pUnkMarshaler.p)
	END_COM_MAP()

	// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);


	DECLARE_PROTECT_FINAL_CONSTRUCT()
	DECLARE_GET_CONTROLLING_UNKNOWN()
	/*
	HRESULT FinalConstruct()
	{
		return CoCreateFreeThreadedMarshaler(
			GetControllingUnknown(), &m_pUnkMarshaler.p);
	}

	void FinalRelease()
	{
		m_pUnkMarshaler.Release();
	}

	CComPtr<IUnknown> m_pUnkMarshaler;
	*/
	
	// ICOMServer implementation
	//
	virtual HRESULT __stdcall NewThread(/*in,optional*/VARIANT script,/*in,optional*/VARIANT params,/*in,optional*/VARIANT title,/*out*/DWORD* ThreadID);
	virtual HRESULT __stdcall ahkPause(/*in*/ VARIANT aThreadID,/*in,optional*/VARIANT aChangeTo,/*out*/VARIANT_BOOL* paused);
	virtual HRESULT __stdcall ahkReady(/*in*/ VARIANT aThreadID,/*out*/VARIANT_BOOL* ready);
	virtual HRESULT __stdcall ahkFindLabel(/*in*/ VARIANT aThreadID,/*[in]*/ VARIANT aLabelName,/*[out, retval]*/ UINT_PTR* pLabel);
	virtual HRESULT __stdcall ahkGetVar(/*in*/ VARIANT aThreadID,/*[in]*/ VARIANT name,/*[in,optional]*/ VARIANT getVar,/*[out, retval]*/ VARIANT* returnVal);
	virtual HRESULT __stdcall ahkAssign(/*in*/ VARIANT aThreadID,/*[in]*/ VARIANT name,/*[in,optional]*/ VARIANT value,/*[out, retval]*/ VARIANT_BOOL* success);
	virtual HRESULT __stdcall ahkExecuteLine(/*in*/ VARIANT aThreadID,/*[in,optional]*/ VARIANT line,/*[in,optional]*/ VARIANT aMode,/*[in,optional]*/ VARIANT wait,/*[out, retval]*/ UINT_PTR* pLine);
	virtual HRESULT __stdcall ahkLabel(/*in*/ VARIANT aThreadID,/*[in]*/ VARIANT aLabelName,/*[in,optional]*/ VARIANT nowait,/*[out, retval]*/ VARIANT_BOOL* success);
	virtual HRESULT __stdcall ahkFindFunc(/*in*/ VARIANT aThreadID,/*[in]*/ VARIANT FuncName,/*[out, retval]*/ UINT_PTR* pFunc);
	virtual HRESULT __stdcall ahkFunction(/*in*/ VARIANT aThreadID,/*[in]*/ VARIANT FuncName,/*[in,optional]*/ VARIANT param1,/*[in,optional]*/ VARIANT param2,/*[in,optional]*/ VARIANT param3,/*[in,optional]*/ VARIANT param4,/*[in,optional]*/ VARIANT param5,/*[in,optional]*/ VARIANT param6,/*[in,optional]*/ VARIANT param7,/*[in,optional]*/ VARIANT param8,/*[in,optional]*/ VARIANT param9,/*[in,optional]*/ VARIANT param10,/*[out, retval]*/ VARIANT* returnVal);
	virtual HRESULT __stdcall ahkPostFunction(/*in*/ VARIANT aThreadID,/*[in]*/ VARIANT FuncName,/*[in,optional]*/ VARIANT param1,/*[in,optional]*/ VARIANT param2,/*[in,optional]*/ VARIANT param3,/*[in,optional]*/ VARIANT param4,/*[in,optional]*/ VARIANT param5,/*[in,optional]*/ VARIANT param6,/*[in,optional]*/ VARIANT param7,/*[in,optional]*/ VARIANT param8,/*[in,optional]*/ VARIANT param9,/*[in,optional]*/ VARIANT param10,/*[out, retval]*/ VARIANT* returnVal);
	virtual HRESULT __stdcall addScript(/*in*/ VARIANT aThreadID,/*[in]*/ VARIANT script,/*[in,optional]*/ VARIANT waitexecute,/*[out, retval]*/ UINT_PTR* success);
	virtual HRESULT __stdcall ahkExec(/*in*/ VARIANT aThreadID,/*[in]*/ VARIANT script,  /*[out, retval]*/ VARIANT_BOOL* success);
};

OBJECT_ENTRY_AUTO(__uuidof(AuotHotkey2), AuotHotkey2)
#endif