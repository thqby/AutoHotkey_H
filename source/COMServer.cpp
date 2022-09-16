// COMServer.cpp: Implementierung von AuotHotkey2
#include "stdafx.h"
#ifdef _USRDLL
#include "COMServer.h"
#include "globaldata.h" // for access to many global vars
#include "script_com.h"
#include <process.h>
#include "exports.h"


// AuotHotkey2

STDMETHODIMP AuotHotkey2::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* const arr[] =
	{
		&IID_ICOMServer
	};

	for (int i = 0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i], riid))
			return S_OK;
	}
	return S_FALSE;
}


//
// AutoHotkey2 functions
//

LPTSTR Variant2T(VARIANT var, LPTSTR buf)
{
	USES_CONVERSION;
	if (var.vt == VT_BYREF + VT_VARIANT)
		var = *var.pvarVal;
	if (var.vt == VT_ERROR)
		return _T("");
	else if (var.vt == VT_BSTR)
		return OLE2T(var.bstrVal);
	else if (var.vt == VT_I2 || var.vt == VT_I4 || var.vt == VT_I8)
#ifdef _WIN64
		return _ui64tot(var.uintVal, buf, 10);
#else
		return _ultot(var.uintVal, buf, 10);
#endif
	return _T("");
}

unsigned int Variant2I(VARIANT var)
{
	USES_CONVERSION;
	if (var.vt == VT_BYREF + VT_VARIANT)
		var = *var.pvarVal;
	if (var.vt == VT_ERROR)
		return 0;
	else if (var.vt == VT_BSTR)
		return ATOI(OLE2T(var.bstrVal));
	else //if (var.vt==VT_I2 || var.vt==VT_I4 || var.vt==VT_I8)
		return var.uintVal;
}

HRESULT __stdcall AuotHotkey2::NewThread(/*in,optional*/VARIANT script,/*in,optional*/VARIANT params,/*in,optional*/VARIANT title,/*out*/DWORD* ThreadID)
{
	USES_CONVERSION;
	TCHAR buf1[MAX_INTEGER_SIZE], buf2[MAX_INTEGER_SIZE], buf3[MAX_INTEGER_SIZE];
	if (ThreadID == NULL)
		return DISP_E_BADPARAMCOUNT;
	*ThreadID = (unsigned int)::NewThread(script.vt == VT_BSTR ? OLE2T(script.bstrVal) : Variant2T(script, buf1)
		, params.vt == VT_BSTR ? OLE2T(params.bstrVal) : Variant2T(params, buf2)
		, title.vt == VT_ERROR ? NULL : title.vt == VT_BSTR ? OLE2T(title.bstrVal) : Variant2T(title, buf3));
	return S_OK;
}
HRESULT __stdcall AuotHotkey2::ahkPause(/*in*/ VARIANT aThreadID,/*in,optional*/VARIANT aChangeTo,/*out*/VARIANT_BOOL* paused)
{
	USES_CONVERSION;
	if (paused == NULL)
		return DISP_E_BADPARAMCOUNT;
	DWORD ThreadID = Variant2I(aThreadID);
	if (ThreadID == 0)
		return E_INVALIDARG;
	TCHAR buf[MAX_INTEGER_SIZE];
	*paused = ::ahkPause(aChangeTo.vt == VT_BSTR ? OLE2T(aChangeTo.bstrVal) : Variant2T(aChangeTo, buf), ThreadID) ? VARIANT_TRUE : VARIANT_FALSE;
	return S_OK;
}
HRESULT __stdcall AuotHotkey2::ahkReady(/*in*/ VARIANT aThreadID,/*out*/VARIANT_BOOL* ready)
{
	if (ready == NULL)
		return DISP_E_BADPARAMCOUNT;
	DWORD ThreadID = Variant2I(aThreadID);
	if (ThreadID == 0)
		return E_INVALIDARG;
	*ready = ::ahkReady(ThreadID) ? VARIANT_TRUE : VARIANT_FALSE;
	return S_OK;
}
HRESULT __stdcall AuotHotkey2::ahkFindLabel(/*in*/ VARIANT aThreadID,/*in*/VARIANT aLabelName,/*out*/UINT_PTR* aLabelPointer)
{
	USES_CONVERSION;
	if (aLabelPointer == NULL)
		return DISP_E_BADPARAMCOUNT;
	DWORD ThreadID = Variant2I(aThreadID);
	if (ThreadID == 0)
		return E_INVALIDARG;
	TCHAR buf[MAX_INTEGER_SIZE];
	*aLabelPointer = ::ahkFindLabel(aLabelName.vt == VT_BSTR ? OLE2T(aLabelName.bstrVal) : Variant2T(aLabelName, buf), ThreadID);
	return S_OK;
}
HRESULT __stdcall AuotHotkey2::ahkGetVar(/*in*/ VARIANT aThreadID,/*in*/VARIANT name,/*[in,optional]*/ VARIANT getVar,/*out*/VARIANT* result)
{
	USES_CONVERSION;
	if (result == NULL)
		return DISP_E_BADPARAMCOUNT;
	DWORD ThreadID = Variant2I(aThreadID);
	AutoTLS atls;
	if (ThreadID == 0 || !atls.Enter(ThreadID))
		return E_INVALIDARG;

	TCHAR buf[MAX_INTEGER_SIZE];
	Var* var;
	ResultToken aToken;

	var = g_script->FindGlobalVar(name.vt == VT_BSTR ? OLE2T(name.bstrVal) : Variant2T(name, buf));
	if (!var)
		return DISP_E_UNKNOWNNAME;
	aToken.InitResult(buf);
	if (var->mType == VAR_VIRTUAL) {
		aToken.symbol = SYM_INTEGER;
		var->Get(aToken);
	}
	else
		var->ToToken(aToken);
	VariantInit(result);
	if (aToken.symbol == SYM_OBJECT) {
		ComObject* comobj;
		IUnknown* obj = aToken.object;
		bool iunk = false;
		LPSTREAM stream;
		HWND hwnd = g_hWnd;
		if (comobj = dynamic_cast<ComObject*>(obj)) {
			if (comobj->mVarType == VT_DISPATCH || comobj->mVarType == VT_UNKNOWN)
				obj = comobj->mUnknown, iunk = comobj->mVarType == VT_UNKNOWN;
			else if (!(comobj->mVarType & ~VT_TYPEMASK))
				comobj->ToVariant(*result), obj = nullptr;
		}
		else if (dynamic_cast<VarRef*>(obj)) {
			aToken.Free();
			return DISP_E_BADVARTYPE;
		}
		if (obj) {
			atls.~AutoTLS();
			if ((stream = (LPSTREAM)SendMessage(hwnd, AHK_THREADVAR, (WPARAM)obj, !iunk)) && SUCCEEDED(CoGetInterfaceAndReleaseStream(stream, iunk ? IID_IUnknown : IID_IDispatch, (LPVOID*)&obj))) {
				result->vt = iunk ? VT_UNKNOWN : VT_DISPATCH;
				result->punkVal = obj;
			}
			else {
				auto err = GetLastError();
				if (stream && !iunk && SUCCEEDED(CoGetInterfaceAndReleaseStream(stream, IID_IUnknown, (LPVOID*)&obj))) {
					result->vt = VT_UNKNOWN;
					result->punkVal = obj;
				}
				else {
					result->vt = VT_ERROR;
					result->scode = err;
					CoReleaseMarshalData(stream), stream->Release();
				}
			}
		}
	}
	else
		TokenToVariant(aToken, *result);
	aToken.Free();
	return S_OK;
}
HRESULT __stdcall AuotHotkey2::ahkAssign(/*in*/ VARIANT aThreadID,/*in*/VARIANT name, /*in*/VARIANT value,/*out*/VARIANT_BOOL* success)
{
	USES_CONVERSION;
	if (success == NULL)
		return DISP_E_BADPARAMCOUNT;
	DWORD ThreadID = Variant2I(aThreadID);
	AutoTLS atls;
	if (ThreadID == 0 || !atls.Enter(ThreadID))
		return E_INVALIDARG;
	TCHAR namebuf[MAX_INTEGER_SIZE];
	ResultToken val;
	LPSTREAM stream = nullptr;
	Var* var;
	HWND hwnd = g_hWnd;
	if (!(var = g_script->FindGlobalVar(name.vt == VT_BSTR ? OLE2T(name.bstrVal) : Variant2T(name, namebuf))))
		return DISP_E_UNKNOWNNAME;
	val.InitResult(_T(""));
	if ((value.vt == VT_DISPATCH && !SUCCEEDED(CoMarshalInterThreadInterfaceInStream(IID_IDispatch, value.punkVal, &stream)))
		|| (value.vt == VT_UNKNOWN && !SUCCEEDED(CoMarshalInterThreadInterfaceInStream(IID_IUnknown, value.punkVal, &stream))))
		return E_FAIL;
	if (stream)
		val.value_int64 = (__int64)stream, val.symbol = SYM_STREAM;
	else
		VariantToToken(value, val);
	atls.~AutoTLS();
	*success = SendMessage(hwnd, AHK_THREADVAR, (WPARAM)&val, (LPARAM)var) ? VARIANT_TRUE : VARIANT_FALSE;
	val.Free();
	return S_OK;
}
HRESULT __stdcall AuotHotkey2::ahkExecuteLine(/*in*/ VARIANT aThreadID,/*[in,optional]*/ VARIANT line,/*[in,optional]*/ VARIANT aMode,/*[in,optional]*/ VARIANT wait,/*[out, retval]*/ UINT_PTR* pLine)
{
	if (pLine == NULL)
		return DISP_E_BADPARAMCOUNT;
	DWORD ThreadID = Variant2I(aThreadID);
	if (ThreadID == 0)
		return E_INVALIDARG;
	*pLine = ::ahkExecuteLine(Variant2I(line), Variant2I(aMode), Variant2I(wait), ThreadID);
	return S_OK;
}
HRESULT __stdcall AuotHotkey2::ahkLabel(/*in*/ VARIANT aThreadID,/*[in]*/ VARIANT aLabelName,/*[in,optional]*/ VARIANT nowait,/*[out, retval]*/ VARIANT_BOOL* success)
{
	USES_CONVERSION;
	if (success == NULL)
		return DISP_E_BADPARAMCOUNT;
	DWORD ThreadID = Variant2I(aThreadID);
	if (ThreadID == 0)
		return E_INVALIDARG;
	TCHAR buf[MAX_INTEGER_SIZE];
	*success = ::ahkLabel(aLabelName.vt == VT_BSTR ? OLE2T(aLabelName.bstrVal) : Variant2T(aLabelName, buf), Variant2I(nowait), ThreadID) ? VARIANT_TRUE : VARIANT_FALSE;
	return S_OK;
}
HRESULT __stdcall AuotHotkey2::ahkFindFunc(/*in*/ VARIANT aThreadID,/*[in]*/ VARIANT FuncName,/*[out, retval]*/ UINT_PTR* pFunc)
{
	USES_CONVERSION;
	if (pFunc == NULL)
		return DISP_E_BADPARAMCOUNT;
	DWORD ThreadID = Variant2I(aThreadID);
	if (ThreadID == 0)
		return E_INVALIDARG;
	TCHAR buf[MAX_INTEGER_SIZE];
	*pFunc = ::ahkFindFunc(FuncName.vt == VT_BSTR ? OLE2T(FuncName.bstrVal) : Variant2T(FuncName, buf), ThreadID);
	return S_OK;
}

HRESULT ahkFunctionVariant(LPTSTR func, VARIANT& ret, VARIANT param1,/*[in,optional]*/ VARIANT param2,/*[in,optional]*/ VARIANT param3,/*[in,optional]*/ VARIANT param4,/*[in,optional]*/ VARIANT param5,/*[in,optional]*/ VARIANT param6,/*[in,optional]*/ VARIANT param7,/*[in,optional]*/ VARIANT param8,/*[in,optional]*/ VARIANT param9,/*[in,optional]*/ VARIANT param10, DWORD aThreadID, int sendOrPost)
{
	AutoTLS atls;
	if (!atls.Enter(aThreadID))
		return E_INVALIDARG;
	Var* aVar = g_script->FindGlobalVar(func);
	Func* aFunc = aVar ? dynamic_cast<Func*>(aVar->ToObject()) : nullptr;
	ret.vt = VT_NULL;
	if (aFunc)
	{
		VARIANT* variants[10] = { &param1,&param2,&param3,&param4,&param5,&param6,&param7,&param8,&param9,&param10 };
		int aParamsCount = 10;
		HWND hwnd = g_hWnd;
		for (; aParamsCount > 0; aParamsCount--)
			if (variants[aParamsCount - 1]->vt != VT_ERROR)
				break;
		if (aParamsCount < aFunc->mMinParams)
			return DISP_E_BADPARAMCOUNT;
		aParamsCount = aFunc->mParamCount < aParamsCount && !aFunc->mIsVariadic ? aFunc->mParamCount : aParamsCount;
		if (sendOrPost) {
			atls.~AutoTLS();
			IDispatch* pdisp;
			if (auto stream = (LPSTREAM)SendMessage(hwnd, AHK_THREADVAR, (WPARAM)aFunc, 1)) {
				if (SUCCEEDED(CoGetInterfaceAndReleaseStream(stream, IID_IDispatch, (void**)&pdisp))) {
					//VARIANTARG rgvarg[10];
					DISPPARAMS dispparams = { &param1, NULL, (UINT)aParamsCount, 0 };
					//for (int i = 0; i < aParamsCount; i++)
					//	memcpy(&rgvarg[i], variants[i], sizeof(VARIANT));
					pdisp->Invoke(DISPID_VALUE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dispparams, &ret, NULL, NULL);
					pdisp->Release();
					return S_OK;
				}
				else
					CoReleaseMarshalData(stream), stream->Release();
			}
			return E_FAIL;
		}
		else {
			atls.~AutoTLS();
			Promise* promise = new Promise(aFunc, !!atls.curr_teb);
			promise->Init(aParamsCount);
			if (g) promise->mPriority = g->Priority;
			LPSTREAM stream;
			for (int i = 0; i < aParamsCount; i++) {
				auto& aToken = promise->mParamToken[i];
				if (variants[i]->vt == VT_BSTR) {
					aToken.marker = _T("");
					aToken.marker_length = 0;
					if (size_t len = SysStringLen(variants[i]->bstrVal)) {
#ifdef UNICODE
						aToken.Malloc(variants[i]->bstrVal, len);
#else
						CStringCharFromWChar buf(variants[i]->bstrVal, len);
						len = buf.GetLength();
						if (aToken.mem_to_free = buf.DetachBuffer())
						{
							aToken.marker = aToken.mem_to_free;
							aToken.marker_length = len;
						}
#endif
					}
				}
				else if (variants[i]->vt == VT_DISPATCH || variants[i]->vt == VT_UNKNOWN) {
					if (SUCCEEDED(CoMarshalInterThreadInterfaceInStream(variants[i]->vt == VT_DISPATCH ? IID_IDispatch : IID_IUnknown, variants[i]->pdispVal, &stream)))
						aToken.value_int64 = (__int64)stream, aToken.symbol = SYM_STREAM;
					else aToken.symbol = SYM_MISSING;
				}
				else {
					VariantToToken(*variants[i], aToken);
					if (aToken.symbol == SYM_OBJECT && SUCCEEDED(CoMarshalInterThreadInterfaceInStream(IID_IDispatch, aToken.object, &stream)))
						aToken.object->Release(), aToken.value_int64 = (__int64)stream, aToken.symbol = SYM_STREAM;
				}
			}
			ret.vt = VT_BOOL;
			if (!(ret.boolVal = PostMessage(hwnd, AHK_EXECUTE_FUNCTION, 0, (LPARAM)promise)))
				promise->Release();
			return S_OK;
		}
	}
	else
		return DISP_E_UNKNOWNNAME;
}

HRESULT __stdcall AuotHotkey2::ahkFunction(/*in*/ VARIANT aThreadID,/*[in]*/ VARIANT FuncName,/*[in,optional]*/ VARIANT param1,/*[in,optional]*/ VARIANT param2,/*[in,optional]*/ VARIANT param3,/*[in,optional]*/ VARIANT param4,/*[in,optional]*/ VARIANT param5,/*[in,optional]*/ VARIANT param6,/*[in,optional]*/ VARIANT param7,/*[in,optional]*/ VARIANT param8,/*[in,optional]*/ VARIANT param9,/*[in,optional]*/ VARIANT param10,/*[out, retval]*/ VARIANT* returnVal)
{
	USES_CONVERSION;
	if (returnVal == NULL)
		return DISP_E_BADPARAMCOUNT;
	DWORD ThreadID = Variant2I(aThreadID);
	if (ThreadID == 0)
		return E_INVALIDARG;
	TCHAR buf[MAX_INTEGER_SIZE];
	return ahkFunctionVariant(FuncName.vt == VT_BSTR ? OLE2T(FuncName.bstrVal) : Variant2T(FuncName, buf), *returnVal
		, param1, param2, param3, param4, param5, param6, param7, param8, param9, param10, ThreadID, 1);
}
HRESULT __stdcall AuotHotkey2::ahkPostFunction(/*in*/ VARIANT aThreadID,/*[in]*/ VARIANT FuncName,/*[in,optional]*/ VARIANT param1,/*[in,optional]*/ VARIANT param2,/*[in,optional]*/ VARIANT param3,/*[in,optional]*/ VARIANT param4,/*[in,optional]*/ VARIANT param5,/*[in,optional]*/ VARIANT param6,/*[in,optional]*/ VARIANT param7,/*[in,optional]*/ VARIANT param8,/*[in,optional]*/ VARIANT param9,/*[in,optional]*/ VARIANT param10,/*[out, retval]*/ VARIANT* returnVal)
{
	USES_CONVERSION;
	if (returnVal == NULL)
		return DISP_E_BADPARAMCOUNT;
	DWORD ThreadID = Variant2I(aThreadID);
	if (ThreadID == 0)
		return E_INVALIDARG;
	TCHAR buf[MAX_INTEGER_SIZE];
	return ahkFunctionVariant(FuncName.vt == VT_BSTR ? OLE2T(FuncName.bstrVal) : Variant2T(FuncName, buf), *returnVal
		, param1, param2, param3, param4, param5, param6, param7, param8, param9, param10, ThreadID, 0);
}
HRESULT __stdcall AuotHotkey2::addScript(/*in*/ VARIANT aThreadID,/*[in]*/ VARIANT script,/*[in,optional]*/ VARIANT waitexecute,/*[out, retval]*/ UINT_PTR* success)
{
	USES_CONVERSION;
	if (success == NULL)
		return DISP_E_BADPARAMCOUNT;
	DWORD ThreadID = Variant2I(aThreadID);
	if (ThreadID == 0)
		return E_INVALIDARG;
	TCHAR buf[MAX_INTEGER_SIZE];
	*success = ::addScript(script.vt == VT_BSTR ? OLE2T(script.bstrVal) : Variant2T(script, buf), Variant2I(waitexecute), ThreadID);
	return S_OK;
}
HRESULT __stdcall AuotHotkey2::ahkExec(/*in*/ VARIANT aThreadID,/*[in]*/ VARIANT script,/*[out, retval]*/ VARIANT_BOOL* success)
{
	USES_CONVERSION;
	if (success == NULL)
		return DISP_E_BADPARAMCOUNT;
	DWORD ThreadID = Variant2I(aThreadID);
	if (ThreadID == 0)
		return E_INVALIDARG;
	TCHAR buf[MAX_INTEGER_SIZE];
	*success = ::ahkExec(script.vt == VT_BSTR ? OLE2T(script.bstrVal) : Variant2T(script, buf), ThreadID) ? VARIANT_TRUE : VARIANT_FALSE;
	return S_OK;
}
#endif