#include "stdafx.h" // pre-compiled headers
#include "globaldata.h" // for access to many global vars
#include "application.h" // for InitNewThread()
#include "script_com.h"
#include "exports.h"
#include <process.h>  // NewThread
#include "LiteZip.h"  // crc32

// Following macros are used in addScript and ahkExec
// HotExpr code from LoadFromFile, Hotkeys need to be toggled to get activated
#define FINALIZE_HOTKEYS \
	if (Hotkey::sHotkeyCount > HotkeyCount)\
	{\
		Hotstring::SuspendAll(!g_IsSuspended);\
		Hotkey::ManifestAllHotkeysHotstringsHooks();\
		Hotstring::SuspendAll(g_IsSuspended);\
		Hotkey::ManifestAllHotkeysHotstringsHooks();\
	}
#define RESTORE_IF_EXPR \
	if (g_FirstHotExpr)\
	{\
		if (aLastHotExpr)\
			aLastHotExpr->NextExpr = g_FirstHotExpr;\
		if (aFirstHotExpr)\
			g_FirstHotExpr = aFirstHotExpr;\
	}\
	else\
	{\
		g_FirstHotExpr = aFirstHotExpr;\
		g_LastHotExpr = aLastHotExpr;\
	}
// AutoHotkey needs to be running at this point
#define BACKUP_G_SCRIPT \
	int aCurrFileIndex = g_script->mCurrFileIndex, aCombinedLineNumber = g_script->mCombinedLineNumber;\
	bool aNextLineIsFunctionBody = g_script->mNextLineIsFunctionBody;\
	Line *aFirstLine = g_script->mFirstLine,*aLastLine = g_script->mLastLine,*aCurrLine = g_script->mCurrLine;\
	UserFunc *aCurrFunc  = g->CurrentFunc, *aCurrMacro  = g->CurrentMacro;\
	int aClassObjectCount = g_script->mClassObjectCount;\
	g_script->mClassObjectCount = NULL;g_script->mFirstLine = NULL;g_script->mLastLine = NULL;g->CurrentFunc = NULL, g->CurrentMacro = NULL;\
	g_script->mNextLineIsFunctionBody = false;g_script->mIsReadyToExecute = false;\
	WarnMode aWarnMode = g_Warn_VarUnset;\
	g_Warn_VarUnset = WARNMODE_OFF;

#define RESTORE_G_SCRIPT \
	g_script->mFirstLine = aFirstLine;\
	g_script->mLastLine = aLastLine;\
	g_script->mLastLine->mNextLine = NULL;\
	g_script->mCurrLine = aCurrLine;\
	g_script->mClassObjectCount = aClassObjectCount + g_script->mClassObjectCount;\
	g_script->mCurrFileIndex = aCurrFileIndex;\
	g_script->mNextLineIsFunctionBody = aNextLineIsFunctionBody;\
	g_script->mCombinedLineNumber = aCombinedLineNumber;\
	g_Warn_VarUnset = aWarnMode;


void callFuncDll(FuncAndToken * aToken, bool throwerr)
{
	IObject* func = aToken->mFunc;
	TCHAR buf[MAX_NUMBER_SIZE];
	aToken->mResult->InitResult(buf);
	if (g_nThreads >= g_MaxThreadsTotal) {
		aToken->mResult->SetResult(FAIL);
		return;
	}

	// See MsgSleep() for comments about the following section.
	InitNewThread(0, false, true);
	g->ExcptMode = EXCPTMODE_CATCH;
	auto result = func->Invoke(*aToken->mResult, IT_CALL, nullptr, ExprTokenType(func), aToken->mParam, aToken->mParamCount);
	if (g->ThrownToken) {
		if (throwerr)
			g_script->UnhandledException(NULL);
		else if (g->ThrownToken->symbol == SYM_STRING && !g->ThrownToken->mem_to_free)
			aToken->mResult->Malloc(g->ThrownToken->marker, g->ThrownToken->marker_length);
		else {
			aToken->mResult->CopyValueFrom(*g->ThrownToken);
			aToken->mResult->mem_to_free = g->ThrownToken->mem_to_free;
			g->ThrownToken->symbol = SYM_INTEGER;
			g->ThrownToken->mem_to_free = nullptr;
		}
	}
	ResumeUnderlyingThread();
	return;
}

void callPromise(Promise * promise, HWND replyHwnd) {
	if (replyHwnd == (HWND)-1) {
		if (promise->mResult.symbol == SYM_STREAM) {
			IUnknown* obj;
			IObject* pobj;
			LPSTREAM stream;
			if (stream = (LPSTREAM)promise->mResult.value_int64) {
				if (SUCCEEDED(CoGetInterfaceAndReleaseStream(stream, IID_IUnknown, (LPVOID*)&obj))) {
					if (SUCCEEDED(obj->QueryInterface(IID_IObjectComCompatible, (LPVOID*)&pobj)))
						obj->Release();
					else if (SUCCEEDED(obj->QueryInterface(IID_IDispatch, (LPVOID*)&pobj)) || SUCCEEDED(obj->QueryInterface(IID__Object, (LPVOID*)&pobj)))
						pobj = new ComObject(pobj), obj->Release();
					else pobj = new ComObject((__int64)obj, VT_UNKNOWN);
					promise->mResult.SetValue(pobj);
					SleepEx(5, true);
				}
				else promise->mResult.SetValue(_T("")), CoReleaseMarshalData(stream), stream->Release();
			}
			else promise->mResult.SetValue(_T(""));
		}
		if (IObject* callback = (promise->mState & 2) ? promise->mError : promise->mComplete) {
			ExprTokenType* param[] = { &promise->mResult };
			ResultToken result;
			TCHAR buf[MAX_INTEGER_LENGTH];
			promise->mState |= 4;
			result.InitResult(buf);
			callback->Invoke(result, IT_CALL, nullptr, ExprTokenType(callback), param, 1);
			result.Free();
		}
	}
	else {
		promise->UnMarshal();
		FuncAndToken token = { promise->mFunc,&promise->mResult,(ExprTokenType**)promise->mParam,promise->mParamCount };
		callFuncDll(&token, false);
		promise->FreeParams();
		if (replyHwnd) {
			bool fail = promise->mResult.Result() == FAIL;
			EnterCriticalSection(&promise->mCritical);
			if (fail)
				promise->mState = 2;
			else
				promise->mState = 1;
			if (promise->mMarshal && promise->mResult.symbol == SYM_OBJECT) {
				auto obj = promise->mResult.object;
				MarshalObjectToToken(obj, promise->mResult);
				obj->Release();
			}
			else if (promise->mResult.symbol == SYM_STRING && !promise->mResult.mem_to_free)
				promise->mResult.Malloc(promise->mResult.marker, promise->mResult.marker_length);
			LeaveCriticalSection(&promise->mCritical);
			if (PostMessage(replyHwnd, AHK_EXECUTE_FUNCTION, -1, (LPARAM)promise))
				return;
		}
		promise->FreeResult();
	}
	promise->Release();
}

EXPORT int ahkReady(DWORD aThreadID) // HotKeyIt check if dll is ready to execute
{
	AutoTLS atls;
	return (int)(atls.Enter(aThreadID) && g_script && !g_Reloading && g_script->mIsReadyToExecute);
}

EXPORT int ahkPause(LPTSTR aChangeTo, DWORD aThreadID) //Change pause state of a running script
{
	AutoTLS atls;
	if (atls.Enter(aThreadID)) {
		if (!g->IsPaused && ((size_t)aChangeTo == 1 || *aChangeTo == '1' || ((*aChangeTo == 'O' || *aChangeTo == 'o') && (*(aChangeTo + 1) == 'N' || *(aChangeTo + 1) == 'n'))))
		{
			Hotkey::ResetRunAgainAfterFinished();
			g->IsPaused = true;
			++g_nPausedThreads; // For this purpose the idle thread is counted as a paused thread.
			g_script->UpdateTrayIcon();
		}
		else if (g->IsPaused)
		{
			g->IsPaused = false;
			--g_nPausedThreads; // For this purpose the idle thread is counted as a paused thread.
			g_script->UpdateTrayIcon();
		}
		return (UINT_PTR)g->IsPaused;
	}
	return 0;
}

EXPORT UINT_PTR ahkFindFunc(LPTSTR aFuncName, DWORD aThreadID) {
	AutoTLS atls;
	return atls.Enter(aThreadID) ? (UINT_PTR)g_script->FindGlobalFunc(aFuncName) : 0;
}

EXPORT UINT_PTR ahkFindLabel(LPTSTR aLabelName, DWORD aThreadID) {
	AutoTLS atls;
	if (atls.Enter(aThreadID))
		for (Label* label = g_script->mFirstLabel; label; label = label->mNextLabel)
			if (!_tcsicmp(label->mName, aLabelName))
				return (UINT_PTR)label;
	return 0;
}

// Naveen: v1. ahkGetVar()
EXPORT LPTSTR ahkGetVar(LPTSTR name, int getVar, DWORD aThreadID)
{
	AutoTLS atls;
	if (atls.Enter(aThreadID)) {
		if (auto ahkvar = g_script->FindGlobalVar(name)) {
			if (getVar) {
				if (ahkvar->mType == VAR_VIRTUAL)
					return NULL;
				return (LPTSTR)(UINT_PTR)ahkvar;
			}
			else {
				TCHAR buf[MAX_NUMBER_SIZE];
				ExprTokenType exp = { _T("") };
				if (!ahkvar->IsObject())
					ahkvar->ToToken(exp);
				auto str = TokenToString(exp, buf);
				return _tcsdup(str);
			}
		}
	}
	return NULL;
}

EXPORT int ahkAssign(LPTSTR name, LPTSTR value, DWORD aThreadID)
{
	AutoTLS atls;
	if (atls.Enter(aThreadID)) {
		HWND hwnd = g_hWnd;
		auto ahkvar = g_script->FindGlobalVar(name);
		if (!ahkvar || ahkvar->IsReadOnly())
			return 0;
		if (!value)
			value = _T("");
		if (atls.tls) {
			ExprTokenType token(value);
			atls.~AutoTLS();
			return (int)SendMessage(hwnd, AHK_THREADVAR, (WPARAM)&token, (LPARAM)ahkvar);
		}
		return ahkvar->Assign(value) == OK;
	}
	return 0;
}

//HotKeyIt ahkExecuteLine()
EXPORT UINT_PTR ahkExecuteLine(UINT_PTR line, int aMode, int wait, DWORD aThreadID)
{
	AutoTLS atls;
	if (atls.Enter(aThreadID)) {
		HWND hwnd = g_hWnd;
		if (line == NULL)
			return (UINT_PTR)g_script->mFirstLine;
		atls.~AutoTLS();
		Line* templine = (Line*)line;
		if (aMode)
		{
			if (!hwnd || wait && hwnd == g_hWnd)
				templine->ExecUntil(aMode == ONLY_ONE_LINE ? ONLY_ONE_LINE : aMode == UNTIL_BLOCK_END ? UNTIL_BLOCK_END : UNTIL_RETURN);
			else if (wait)
				SendMessage(hwnd, AHK_EXECUTE, (WPARAM)line, (LPARAM)aMode);
			else
				PostMessage(hwnd, AHK_EXECUTE, (WPARAM)line, (LPARAM)aMode);
		}
		if (templine = templine->mNextLine)
			if (templine->mActionType == ACT_BLOCK_BEGIN && templine->mAttribute)
			{
				for (; !(templine->mActionType == ACT_BLOCK_END && templine->mAttribute);)
					templine = templine->mNextLine;
				templine = templine->mNextLine;
			}
		return (UINT_PTR)templine;
	}
	return 0;
}

EXPORT int ahkLabel(LPTSTR aLabelName, int nowait, DWORD aThreadID) // 0 = wait = default
{
	AutoTLS atls;
	Label* label;
	if (atls.Enter(aThreadID)) {
		HWND hwnd = g_hWnd;
		for (label = g_script->mFirstLabel; label; label = label->mNextLabel)
			if (!_tcsicmp(label->mName, aLabelName))
				break;
		atls.~AutoTLS();
		if (label) {
			if (nowait)
				PostMessage(hwnd, AHK_EXECUTE_LABEL, (WPARAM)label, 0);
			else
				SendMessage(hwnd, AHK_EXECUTE_LABEL, (WPARAM)label, 0);
			return 1;
		}
	}
	return 0;
}

struct VarListBackup {
	VarList& list;
	Var** mItem;
	int mCount;
	VarListBackup(VarList& varlist) : list(varlist), mCount(list.mCount), mItem(nullptr) {
		if (mCount) {
			if (mItem = (Var**)malloc(mCount * sizeof(Var*)))
				memcpy(mItem, list.mItem, mCount * sizeof(Var*));
			else mCount = -1;
		}
	}
	~VarListBackup() {
		if (mCount < 0) return;
		if (mCount < list.mCount) {
			auto vars = list.mItem;
			auto len = list.mCount;
			for (int i = 0, j = 0; i < len; i++) {
				auto var = j < mCount ? mItem[j] : nullptr;
				while (i < len && var != vars[i])
					vars[i]->Free(VAR_ALWAYS_FREE, true), vars[i++] = NULL;
				vars[j++] = var;
			}
			list.mCount = mCount;
		}
		free(mItem), mCount = -1;
	}
};

// HotKeyIt: addScript()
EXPORT UINT_PTR addScript(LPTSTR script, int waitexecute, DWORD aThreadID)
{
	AutoTLS atls;
	int ret = atls.Enter(aThreadID);
	if (!ret)
		return 0;
	else if (ret == 2) {
		ExprTokenType params[3] = { (__int64)addScript, _T("i==siui") };
		ExprTokenType* param[] = { params, params + 1, params + 2 };
		ResultToken result;
		IObject* func = DynaToken::Create(param, 2);
		FuncAndToken token = { func,&result,param,3 };
		auto hwnd = g_hWnd;
		params[0].SetValue(script), params[1].SetValue(waitexecute), params[2].SetValue((__int64)aThreadID);
		result.InitResult(_T(""));
		atls.~AutoTLS();
		SendMessage(hwnd, AHK_EXECUTE_FUNCTION, (WPARAM)&token, 0);
		func->Release();
		return (UINT_PTR)TokenToInt64(result);
	}
	EnterCriticalSection(&g_CriticalTLSCallback);
	int aVarCount = g_script->mVars.mCount;
	int aFuncCount = g_script->mFuncs.mCount;
	int aHotFuncCount = g_script->mHotFuncs.mCount;
	VarListBackup aVarBkp(g_script->mVars);
	if (aVarBkp.mCount == -1) {
		LeaveCriticalSection(&g_CriticalTLSCallback);
		return 0;
	}
	Line* aLastStaticLine = g_script->mLastStaticLine;
	int HotkeyCount = Hotkey::sHotkeyCount;
	HotkeyCriterion* aFirstHotExpr = g_FirstHotExpr, * aLastHotExpr = g_LastHotExpr;
	g_FirstHotExpr = NULL; g_LastHotExpr = NULL;
	int aSourceFileIdx = Line::sSourceFileCount;
	
	// Backup SimpleHeap to restore later
	SimpleHeap* aSimpleHeap = new SimpleHeap();
	SimpleHeap* bkpSimpleHeap = g_SimpleHeap;
	TCHAR tmp[MAX_PATH];
	void* p = nullptr;
	g_SimpleHeap = aSimpleHeap;
	if (!g_script->mEncrypt)
		p = g_SimpleHeap->Alloc(script);
	_stprintf(tmp, _T("*THREAD%u?%p#%zu.AHK"), g_MainThreadID, p, _tcslen(script) * sizeof(TCHAR));
	BACKUP_G_SCRIPT
	if (g_script->LoadFromText(script,tmp) != TRUE)
	{
		RESTORE_IF_EXPR
		g->CurrentFunc = (UserFunc*)aCurrFunc;
		g->CurrentMacro = (UserFunc*)aCurrMacro;
		Line* aTempLine = g_script->mLastLine;
		Line* aExecLine = g_script->mFirstLine;
		if (g_script->mLastStaticLine && g_script->mLastStaticLine != aLastStaticLine && aExecLine) {
			aExecLine->mPrevLine = g_script->mLastStaticLine;
			if (aLastStaticLine)
				aLastStaticLine->mNextLine->mPrevLine = NULL, aLastStaticLine->mNextLine = NULL;
		}
		g_script->mLastStaticLine = aLastStaticLine;
		g_script->mIsReadyToExecute = true;
		RESTORE_G_SCRIPT;
		g_SimpleHeap = bkpSimpleHeap;
		Line::sSourceFileCount = aSourceFileIdx;
		aVarBkp.~VarListBackup();
		int len = g_script->mHotFuncs.mCount;
		if (len > aHotFuncCount) {
			auto& funcs = g_script->mHotFuncs.mItem;
			for (int i = aHotFuncCount; i < len; i++)
				funcs[i]->Release(), funcs[i] = NULL;
			g_script->mHotFuncs.mCount = aHotFuncCount;
		}
		len = g_script->mFuncs.mCount;
		if (len > aFuncCount) {
			auto& funcs = g_script->mFuncs.mItem;
			for (int i = aFuncCount; i < len; i++) {
				if (auto cl = ((UserFunc*)(funcs[i]))->mClass)
					cl->Release();
				funcs[i]->Release(), funcs[i] = NULL;
			}
			g_script->mFuncs.mCount = aFuncCount;
		}
		for (Line* line = aTempLine; line; line = line->mPrevLine) {
			for (int i = 0; i < line->mArgc; ++i) {
				ArgStruct& this_arg = line->mArg[i];
				if (!this_arg.is_expression || !this_arg.postfix)
					continue;
				for (ExprTokenType* token = this_arg.postfix; token->symbol != SYM_INVALID; ++token) {
					if (token->symbol == SYM_OBJECT)
						token->object->Release();
				}
			}
			line->FreeDerefBufIfLarge();
			if (line->mBreakpoint)
				delete line->mBreakpoint;
			delete line;
		}
		// Delete used and restore SimpleHeap
		delete aSimpleHeap;
		LeaveCriticalSection(&g_CriticalTLSCallback);
		return 0;
	}
	FINALIZE_HOTKEYS
	RESTORE_IF_EXPR
	aVarBkp.mCount = aVarBkp.list.mCount;
	g_script->mIsReadyToExecute = true;
	g->CurrentFunc = (UserFunc*)aCurrFunc;
	g->CurrentMacro = (UserFunc*)aCurrMacro;
	g_SimpleHeap = bkpSimpleHeap;
	g_SimpleHeap->Merge(aSimpleHeap);
	Line *aTempLine = g_script->mFirstLine;
	aLastLine->mNextLine = aTempLine;
	aTempLine->mPrevLine = aLastLine;
	aLastLine = g_script->mLastLine;
	RESTORE_G_SCRIPT;
	if (waitexecute) {
		if (waitexecute == 1 || !g_hWnd)
			aTempLine->ExecUntil(UNTIL_RETURN);
		else
			PostMessage(g_hWnd, AHK_EXECUTE, (WPARAM)aTempLine, 1);
	}
	LeaveCriticalSection(&g_CriticalTLSCallback);
	return (UINT_PTR)aTempLine;
}

EXPORT int ahkExec(LPTSTR script, DWORD aThreadID)
{
	AutoTLS atls;
	int ret = atls.Enter(aThreadID);
	if (!ret)
		return 0;
	else if (ret == 2) {
		ExprTokenType params[2] = { (__int64)ahkExec, _T("i==sui") };
		ExprTokenType* param[] = { params, params + 1 };
		ResultToken result;
		IObject* func = DynaToken::Create(param, 2);
		FuncAndToken token = { func,&result,param,2 };
		auto hwnd = g_hWnd;
		params[0].SetValue(script), params[1].SetValue((__int64)aThreadID);
		result.InitResult(_T(""));
		atls.~AutoTLS();
		SendMessage(hwnd, AHK_EXECUTE_FUNCTION, (WPARAM)&token, 0);
		func->Release();
		return (int)TokenToInt64(result);
	}
	EnterCriticalSection(&g_CriticalTLSCallback);
	int aVarCount = g_script->mVars.mCount;
	int aFuncCount = g_script->mFuncs.mCount;
	int aHotFuncCount = g_script->mHotFuncs.mCount;
	VarListBackup aVarBkp(g_script->mVars);
	VarListBackup* fvbkp = nullptr, * fsvbkp = nullptr;
	if (aVarBkp.mCount == -1) {
		LeaveCriticalSection(&g_CriticalTLSCallback);
		return 0;
	}
	if (!aThreadID) {
		auto cur = g->CurrentMacro;
		if (!cur)
			cur = g->CurrentFunc;
		if (cur) {
			fvbkp = new VarListBackup(cur->mVars);
			fsvbkp = new VarListBackup(cur->mStaticVars);
			if (fvbkp->mCount == -1 || fsvbkp->mCount == -1) {
				delete fvbkp;
				delete fsvbkp;
				LeaveCriticalSection(&g_CriticalTLSCallback);
				return 0;
			}
		}
	}
	Line* aLastStaticLine = g_script->mLastStaticLine;
	int HotkeyCount = Hotkey::sHotkeyCount;
	HotkeyCriterion *aFirstHotExpr = g_FirstHotExpr,*aLastHotExpr = g_LastHotExpr;
	g_FirstHotExpr = NULL;g_LastHotExpr = NULL;
	int aSourceFileIdx = Line::sSourceFileCount;
	BACKUP_G_SCRIPT

	// Backup SimpleHeap to restore later
	SimpleHeap *aSimpleHeap = new SimpleHeap();
	SimpleHeap *bkpSimpleHeap = g_SimpleHeap;
	g_SimpleHeap = aSimpleHeap;

	if (!aThreadID)
		g->CurrentMacro = g->CurrentFunc = aCurrMacro ? aCurrMacro : aCurrFunc;

	TCHAR tmp[MAX_PATH];
	_stprintf(tmp, _T("*THREAD%u?%p#%zu.AHK"), g_MainThreadID, g_script->mEncrypt ? nullptr : script, _tcslen(script) * sizeof(TCHAR));
	auto result = g_script->LoadFromText(script, tmp);
	FINALIZE_HOTKEYS
	RESTORE_IF_EXPR
	g->CurrentFunc = (UserFunc*)aCurrFunc;
	g->CurrentMacro = (UserFunc*)aCurrMacro;
	Line *aTempLine = g_script->mLastLine;
	Line *aExecLine = g_script->mFirstLine;
	if (g_script->mLastStaticLine && g_script->mLastStaticLine != aLastStaticLine && aExecLine) {
		aExecLine->mPrevLine = g_script->mLastStaticLine;
		if (aLastStaticLine && aLastStaticLine->mNextLine)
			aLastStaticLine->mNextLine->mPrevLine = NULL, aLastStaticLine->mNextLine = NULL;
	}
	g_script->mLastStaticLine = aLastStaticLine;
	g_script->mIsReadyToExecute = true;
	RESTORE_G_SCRIPT;
	// Restore SimpleHeap so functions will use correct memory
	g_SimpleHeap = bkpSimpleHeap;
	
	if (result == TRUE) {
		g_script->mLastLine->mNextLine = aExecLine;
		aExecLine->ExecUntil(UNTIL_RETURN);
		g_script->mLastLine->mNextLine = NULL;
		g_script->mCurrLine = aCurrLine;
	}
	Line::sSourceFileCount = aSourceFileIdx;
	if (fvbkp) {
		delete fvbkp;
		delete fsvbkp;
	}
	aVarBkp.~VarListBackup();
	int len = g_script->mHotFuncs.mCount;
	if (len > aHotFuncCount) {
		auto& funcs = g_script->mHotFuncs.mItem;
		for (int i = aHotFuncCount; i < len; i++)
			funcs[i]->Release(), funcs[i] = NULL;
		g_script->mHotFuncs.mCount = aHotFuncCount;
	}
	len = g_script->mFuncs.mCount;
	if (len > aFuncCount) {
		auto& funcs = g_script->mFuncs.mItem;
		for (int i = aFuncCount; i < len; i++) {
			if (auto cl = ((UserFunc*)(funcs[i]))->mClass)
				cl->Release();
			funcs[i]->Release(), funcs[i] = NULL;
		}
		g_script->mFuncs.mCount = aFuncCount;
	}
	for (Line* line = aTempLine; line; line = line->mPrevLine) {
		for (int i = 0; i < line->mArgc; ++i) {
			ArgStruct& this_arg = line->mArg[i];
			if (!this_arg.is_expression || !this_arg.postfix)
				continue;
			for (ExprTokenType* token = this_arg.postfix; token->symbol != SYM_INVALID; ++token) {
				if (token->symbol == SYM_OBJECT)
					token->object->Release();
			}
		}
		line->FreeDerefBufIfLarge();
		if (line->mBreakpoint)
			delete line->mBreakpoint;
		delete line;
	}
	// Delete used and restore SimpleHeap
	delete aSimpleHeap;
	LeaveCriticalSection(&g_CriticalTLSCallback);
	return result == TRUE;
}

LPTSTR ahkFunction(LPTSTR func, LPTSTR param1, LPTSTR param2, LPTSTR param3, LPTSTR param4, LPTSTR param5, LPTSTR param6, LPTSTR param7, LPTSTR param8, LPTSTR param9, LPTSTR param10, DWORD aThreadID, int sendOrPost)
{
	AutoTLS atls;
	if (!atls.Enter(aThreadID))
		return NULL;
	Var* aVar = g_script->FindGlobalVar(func);
	Func* aFunc = aVar ? dynamic_cast<Func*>(aVar->ToObject()) : nullptr;
	if (aFunc)
	{
		int aParamsCount = 10;
		LPTSTR* params[10] = { &param1,&param2,&param3,&param4,&param5,&param6,&param7,&param8,&param9,&param10 };
		for (; aParamsCount > 0; aParamsCount--)
			if (*params[aParamsCount - 1])
				break;
		if (aParamsCount < aFunc->mMinParams)
		{
			g_script->ScriptError(ERR_TOO_FEW_PARAMS);
			return NULL;
		}
		Promise* promise = nullptr;
		ExprTokenType paramtoken[10];
		ExprTokenType* param[10] = { 0 };
		HWND hwnd = g_hWnd;
		aParamsCount = aFunc->mParamCount < aParamsCount && !aFunc->mIsVariadic ? aFunc->mParamCount : aParamsCount;
		
		if (sendOrPost) {
			for (int i = 0; aParamsCount > i; i++) {
				param[i] = &paramtoken[i];
				if (*params[i])
					paramtoken[i].SetValue(*params[i]);
				else paramtoken[i].symbol = SYM_MISSING;
			}
			ResultToken result;
			FuncAndToken aFuncAndToken = { aFunc, &result, param, aParamsCount };
			atls.~AutoTLS();
			if (!hwnd || !atls.curr_teb) {
				callFuncDll(&aFuncAndToken, true);
				if (result.symbol == SYM_OBJECT)
					result.object->Release();
			}
			else SendMessage(hwnd, AHK_EXECUTE_FUNCTION, (WPARAM)&aFuncAndToken, 0);
			if (result.Result() == FAIL)
				return NULL;
			TCHAR buf[MAX_INTEGER_LENGTH];
			auto str = TokenToString(result, buf);
			if (!result.mem_to_free)
				str = _tcsdup(str);
			return str;
		}
		else {
			atls.~AutoTLS();
			promise = new Promise(aFunc, false);
			promise->Init(aParamsCount);
			if (g) promise->mPriority = g->Priority;
			for (int i = 0; i < aParamsCount; i++)
				if (*params[i])
					promise->mParam[i]->Malloc(*params[i], -1);
				else promise->mParam[i]->symbol = SYM_MISSING;
			bool ret;
			if (!(ret = PostMessage(hwnd, AHK_EXECUTE_FUNCTION, 0, (LPARAM)promise)))
				promise->Release();
			return (LPTSTR)ret;
		}
	}
	return NULL;
}

EXPORT LPTSTR ahkFunction(LPTSTR func, LPTSTR param1, LPTSTR param2, LPTSTR param3, LPTSTR param4, LPTSTR param5, LPTSTR param6, LPTSTR param7, LPTSTR param8, LPTSTR param9, LPTSTR param10, DWORD aThreadID)
{
	return ahkFunction(func, param1, param2, param3, param4, param5, param6, param7, param8, param9, param10, aThreadID, 1);
}

EXPORT int ahkPostFunction(LPTSTR func, LPTSTR param1, LPTSTR param2, LPTSTR param3, LPTSTR param4, LPTSTR param5, LPTSTR param6, LPTSTR param7, LPTSTR param8, LPTSTR param9, LPTSTR param10, DWORD aThreadID)
{
	return (int)(UINT_PTR)ahkFunction(func, param1, param2, param3, param4, param5, param6, param7, param8, param9, param10, aThreadID, 0);
}

EXPORT DWORD NewThread(LPTSTR aScript, LPTSTR aCmdLine, LPTSTR aTitle) // HotKeyIt check if dll is ready to execute
{
	UINT_PTR data[5] = { 0 };
	if (aScript && *aScript)
		data[1] = (UINT_PTR)aScript, data[0] += _tcsclen(aScript) + 1;
	if (aCmdLine && *aCmdLine)
		data[2] = (UINT_PTR)aCmdLine;
	if (aTitle && *aTitle)
		data[3] = (UINT_PTR)aTitle, data[0] += _tcsclen(aTitle) + 1;
	unsigned int ThreadID;
	TCHAR buf[MAX_INTEGER_LENGTH];
	HANDLE hThread = (HANDLE)_beginthreadex(NULL, 0, (unsigned(__stdcall*)(void*)) & ThreadMain, data, 0, &ThreadID);
	if (hThread)
	{
		CloseHandle(hThread);
		sntprintf(buf, _countof(buf), _T("ahk%d"), ThreadID);
		if (HANDLE hEvent = CreateEvent(NULL, false, false, buf))
		{
			// we need to give the thread also little time to allocate memory to avoid std::bad_alloc that might happen when you run newthread in a loop
			WaitForSingleObject(hEvent, 1000);
			CloseHandle(hEvent);
		}
	}
	return data[4] ? ThreadID : 0;
}

IAhkApi IAhkApi::instance;
thread_local Object* IAhkApi::sObject[(int)ObjectType::MAXINDEX] = {};
thread_local Object* IAhkApi::sFuncPrototype = nullptr;
thread_local void* IAhkApi::sMemToFree = nullptr;
thread_local int IAhkApi::sInit = 0;

IAhkApi::Prototype::~Prototype() { if (OnDispose) OnDispose(); }

Object* IAhkApi::Prototype::New(ExprTokenType* aParam[], int aParamCount) {
	auto size = max(mSize, sizeof(Object));
	auto obj = new (malloc(size)) Object;
	if (!obj) return nullptr;
	obj->SetBase(this);
	if (size -= sizeof(Object))
		memset((char*)obj + sizeof(Object), 0, size);
	FuncResult result;
	return obj->Construct(result, aParam, aParamCount) == OK ? obj : nullptr;
}

IAhkApi* IAhkApi::Initialize() {
	int fg = 0, i = 0;
#ifdef _USRDLL
	if (!g_script) {
		InitializeCriticalSection(&g_CriticalRegExCache);
		InitializeCriticalSection(&g_CriticalTLSCallback);
		g_Debugger = new Debugger();
		g_clip = new Clipboard();
		g_script = new Script();
		g_MsgMonitor = new MsgMonitorList();
		g_SimpleHeap = new SimpleHeap();
		Object::sAnyPrototype = Object::CreateRootPrototypes();
		SetErrorMode(SEM_FAILCRITICALERRORS);
		UpdateWorkingDir();
		g_WorkingDirOrig = g_SimpleHeap->Malloc(const_cast<LPTSTR>(g_WorkingDir.GetString()));
		global_init(*g);
		g_persistent = true;
		g_NoTrayIcon = true;
		g_script->Init(*g, _T("* "), 0, g_hInstance, g_WindowClassMain);
		g_script->mFileName = g_WindowClassMain;
		g_script->LoadFromText(_T(" "));
		g_script->CreateWindows();
		setvbuf(stdout, NULL, _IONBF, 0);
		Hotkey::ManifestAllHotkeysHotstringsHooks();
		g_HSSameLineAction = false;
		g_SuspendExempt = false;
		g_script->mIsReadyToExecute = true;
		fg = 2;
		EnterCriticalSection(&g_Critical);
		for (i = 1; i < MAX_AHK_THREADS; i++)
			if (!IsWindow(g_ahkThreads[i].Hwnd))
				break;
		if (i < MAX_AHK_THREADS)
			g_ahkThreads[i] = { g_hWnd,g_script,((PMYTEB)NtCurrentTeb())->ThreadLocalStoragePointer,g_MainThreadID };
		LeaveCriticalSection(&g_Critical);
		g_script->AutoExecSection();
	}
#endif // _USRDLL
	if (!g_script)
		return nullptr;
	if (!instance.sInit) {
#define GETAHKOBJECT(type) instance.sObject[(int)ObjectType::type] = g_script->FindClass(_T(#type))
		GETAHKOBJECT(Object);
		GETAHKOBJECT(Array);
		GETAHKOBJECT(Buffer);
		GETAHKOBJECT(Map);
		GETAHKOBJECT(ComValue);
		GETAHKOBJECT(ComObjArray);
		GETAHKOBJECT(ComObject);
		GETAHKOBJECT(Worker);
		instance.sObject[(int)ObjectType::File] = g_script->FindGlobalFunc(_T("FileOpen"));
		instance.sObject[(int)ObjectType::UObject] = g_script->FindGlobalFunc(_T("UObject"));
		instance.sObject[(int)ObjectType::UArray] = g_script->FindGlobalFunc(_T("UArray"));
		instance.sObject[(int)ObjectType::UMap] = g_script->FindGlobalFunc(_T("UMap"));
		sFuncPrototype = Object::CreatePrototype(_T("NativeFunc"), Func::sPrototype);
		auto __del = new BuiltInFunc(_T("NativeFunc.Prototype.__Delete"), [](BIF_DECL_PARAMS) {
			if (auto obj = dynamic_cast<NativeFunc*>(TokenToObject(*aParam[0]))) {
				free(sMemToFree);
				sMemToFree = obj;
			}
			else _f_throw_param(0);
			}, 1, 1);
		sFuncPrototype->DefineMethod(_T("__Delete"), __del);
		__del->Release();
		instance.sInit = 1 | fg;
#undef GETAHKOBJECT
	}
	return &instance;
}

void IAhkApi::Finalize() {
	if (sFuncPrototype)
		sFuncPrototype->Release(), sFuncPrototype = nullptr;
	if (sMemToFree)
		free(sMemToFree), sMemToFree = nullptr;
	memset(sObject, 0, sizeof(sObject));
	sInit = 0;
}

STDMETHODIMP IAhkApi::QueryInterface(REFIID riid, void** ppv) { if (riid == IID_IUnknown) { *ppv = this; return S_OK; } return E_NOINTERFACE; }

ULONG STDMETHODCALLTYPE IAhkApi::AddRef() { return 1; }

ULONG STDMETHODCALLTYPE IAhkApi::Release() { return 1; }

void* STDMETHODCALLTYPE IAhkApi::Malloc(size_t aSize) { return malloc(aSize); }

void* STDMETHODCALLTYPE IAhkApi::Realloc(void* aPtr, size_t aSize) { return realloc(aPtr, aSize); }

void STDMETHODCALLTYPE IAhkApi::Free(void* aPtr) { free(aPtr); }

LPTSTR STDMETHODCALLTYPE IAhkApi::TokenToString(ExprTokenType& aToken, TCHAR aBuf[], size_t* aLength) { return ::TokenToString(aToken, aBuf, aLength); }

bool STDMETHODCALLTYPE IAhkApi::TokenToNumber(ExprTokenType& aInput, ExprTokenType& aOutput) { return TokenToDoubleOrInt64(aInput, aOutput) == OK; }

bool STDMETHODCALLTYPE IAhkApi::VarAssign(Var* aVar, ExprTokenType& aToken) { return aVar->Assign(aToken) == OK; }

void STDMETHODCALLTYPE IAhkApi::VarToToken(Var* aVar, ExprTokenType& aToken) { aVar->ToTokenSkipAddRef(aToken); }

void STDMETHODCALLTYPE IAhkApi::VarFree(Var* aVar, bool aExcludeAliasesAndRequireInit) { aVar->Free(VAR_ALWAYS_FREE, aExcludeAliasesAndRequireInit); }

bool STDMETHODCALLTYPE IAhkApi::VariantAssign(Object::Variant& aVariant, ExprTokenType& aValue) { return aVariant.Assign(aValue); }

void STDMETHODCALLTYPE IAhkApi::VariantToToken(Object::Variant& aVariant, ExprTokenType& aToken) { aVariant.ToToken(aToken); }

void STDMETHODCALLTYPE IAhkApi::VariantToToken(VARIANT& aVariant, ResultToken& aToken, bool aRetainVar) { ::VariantToToken(aVariant, aToken, aRetainVar); }

void STDMETHODCALLTYPE IAhkApi::ResultTokenFree(ResultToken& aToken) { aToken.Free(), aToken.InitResult(aToken.buf); }

ResultType STDMETHODCALLTYPE IAhkApi::Error(LPTSTR aErrorText, LPTSTR aExtraInfo, ErrorType aType) {
	Object* proto = nullptr;
#define TEST_TYPE(type) else if (aType == ErrorType::type) proto = ErrorPrototype::type
	if (false) {}
	TEST_TYPE(Error);
	TEST_TYPE(Index);
	TEST_TYPE(Member);
	TEST_TYPE(Property);
	TEST_TYPE(Method);
	TEST_TYPE(Memory);
	TEST_TYPE(OS);
	TEST_TYPE(Target);
	TEST_TYPE(Timeout);
	TEST_TYPE(Type);
	TEST_TYPE(Unset);
	TEST_TYPE(UnsetItem);
	TEST_TYPE(Value);
	TEST_TYPE(ZeroDivision);
#undef TEST_TYPE
	ResultToken result{};
	result.SetResult(OK);
	result.Error(aErrorText, aExtraInfo, proto);
	return result.Result();
}

ResultType STDMETHODCALLTYPE IAhkApi::TypeError(LPTSTR aExpectedType, ExprTokenType& aToken) {
	ResultToken result{};
	result.SetResult(OK);
	result.TypeError(aExpectedType, aToken);
	return result.Result();
}

void* STDMETHODCALLTYPE IAhkApi::GetProcAddress(LPTSTR aDllFileFunc, HMODULE* hmodule_to_free) { return ::GetDllProcAddress(aDllFileFunc, hmodule_to_free); }

void* STDMETHODCALLTYPE IAhkApi::GetProcAddressCrc32(HMODULE aModule, UINT aCRC32, UINT aInitial) {
	static auto RtlComputeCrc32 = (DWORD(WINAPI*)(DWORD dwInitial, BYTE * pData, INT iLen))::GetProcAddress(GetModuleHandleA("ntdll"), "RtlComputeCrc32");
	if (!aModule)
		aModule = GetModuleHandleA("kernel32");
	PIMAGE_DOS_HEADER lpDosHeader = (PIMAGE_DOS_HEADER)aModule;
	PIMAGE_NT_HEADERS lpNtHeader = (PIMAGE_NT_HEADERS)((UINT_PTR)aModule + lpDosHeader->e_lfanew);
	auto& entry = lpNtHeader->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXPORT];
	if (!entry.Size || !entry.VirtualAddress)
		return NULL;
	PIMAGE_EXPORT_DIRECTORY lpExports = (PIMAGE_EXPORT_DIRECTORY)((UINT_PTR)aModule + entry.VirtualAddress);
	PDWORD lpNames = (PDWORD)((UINT_PTR)aModule + lpExports->AddressOfNames);
	PWORD lpOrdinals = (PWORD)((UINT_PTR)aModule + lpExports->AddressOfNameOrdinals);
	PDWORD lpFuncs = (PDWORD)((UINT_PTR)aModule + lpExports->AddressOfFunctions);

	for (DWORD i = 0; i < lpExports->NumberOfNames; i++) {
		char* name = (char*)(lpNames[i] + (UINT_PTR)aModule);
		if (crc32(aInitial, (BYTE*)name, (DWORD)strlen(name)) == aCRC32)
			return (void*)((UINT_PTR)aModule + lpFuncs[lpOrdinals[i]]);
	}
	return NULL;
}

bool STDMETHODCALLTYPE IAhkApi::Script_GetVar(LPTSTR aVarName, ExprTokenType& aValue) {
	auto var = g_script ? g_script->FindGlobalVar(aVarName) : nullptr;
	if (var) {
		var->ToTokenSkipAddRef(aValue);
		return true;
	}
	return false;
}

bool STDMETHODCALLTYPE IAhkApi::Script_SetVar(LPTSTR aVarName, ExprTokenType& aValue) {
	auto var = g_script ? g_script->FindOrAddVar(aVarName, 0, FINDVAR_GLOBAL) : nullptr;
	if (var)
		return var->Assign(aValue) == OK;
	return false;
}

Func* STDMETHODCALLTYPE IAhkApi::Func_New(FuncEntry& aBIF) {
	BuiltInFunc tmp(aBIF);
	auto size = (_tcslen(aBIF.mName) + 1) * sizeof(TCHAR);
	HMODULE mod = nullptr;
	if (tmp.mOutputVars && !GetModuleHandleEx(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, (LPCTSTR)&aBIF, &mod))
		size += MAX_FUNC_OUTPUT_VAR;
	auto func = (BuiltInFunc*)malloc(sizeof(BuiltInFunc) + size);
	if (!func) return nullptr;
	memcpy(func, &tmp, sizeof(BuiltInFunc));
	if (func->mOutputVars && !mod) {
		func->mOutputVars = (UCHAR*)func + sizeof(BuiltInFunc);
		memcpy(func->mOutputVars, aBIF.mOutputVars, MAX_FUNC_OUTPUT_VAR);
		func->mName = (LPCTSTR)(func->mOutputVars + MAX_FUNC_OUTPUT_VAR);
		size -= MAX_FUNC_OUTPUT_VAR;
	}
	else func->mName = (LPCTSTR)((UCHAR*)func + sizeof(BuiltInFunc));
	memcpy((void*)func->mName, aBIF.mName, size);
	tmp.Base()->AddRef();
	func->SetBase(IAhkApi::sFuncPrototype);
	return func;
}

Func* STDMETHODCALLTYPE IAhkApi::Method_New(LPTSTR aFullName, ObjectMember& aMember, Object* aPrototype) {
	BuiltInMethod tmp(aFullName);
	auto size = (_tcslen(aFullName) + 1) * sizeof(TCHAR);
	auto func = (BuiltInMethod*)malloc(sizeof(BuiltInMethod) + size);
	if (!func) return nullptr;
	memcpy(func, &tmp, sizeof(BuiltInMethod));
	func->mName = (LPCTSTR)((char*)func + sizeof(BuiltInMethod));
	memcpy((void*)func->mName, aFullName, size);
	tmp.Base()->AddRef();
	func->SetBase(IAhkApi::sFuncPrototype);
	func->mMID = aMember.id;
	func->mClass = aPrototype;
	func->mBIM = aMember.method;
	func->mMIT = aMember.invokeType;
	func->mIsVariadic = aMember.maxParams == MAXP_VARIADIC;
	if (func->mMIT == IT_CALL) {
		func->mMinParams = aMember.minParams + 1;
		func->mParamCount = func->mIsVariadic ? func->mMinParams : aMember.maxParams + 1;
		if (aPrototype)
			aPrototype->DefineMethod(aMember.name, func);
	}
	else {
		int n = func->mMIT == IT_GET ? 1 : 2;
		func->mMinParams = aMember.minParams + n;
		func->mParamCount = aMember.maxParams + n;
		if (aPrototype) {
			auto prop = aPrototype->DefineProperty(aMember.name);
			prop->MinParams = aMember.minParams;
			prop->MaxParams = aMember.maxParams;
			if (func->mMIT == IT_GET)
				prop->SetGetter(func);
			else prop->SetSetter(func);
		}
	}
	return func;
}

Object* STDMETHODCALLTYPE IAhkApi::Class_New(LPTSTR aClassName, size_t aClassSize, ObjectMember aMember[], int aMemberCount, Prototype*& aPrototype, Object* aBase) {
	static BuiltInFunctionType obj_ctor = [](BIF_DECL_PARAMS) {
		auto proto = static_cast<Prototype*>(aResultToken.func->mData);
		auto size = max(proto->mSize, sizeof(Object));
		Object* obj = new (malloc(size)) Object;
		if (!obj)
			_f_throw_oom;
		if (size -= sizeof(Object))
			memset((char*)obj + sizeof(Object), 0, size);
		obj->SetBase(proto);
		obj->New(aResultToken, aParam, aParamCount);
	};
	TCHAR full_name[MAX_VAR_NAME_LENGTH + 1];
	Prototype* proto = nullptr;
	if (!aPrototype) {
		aPrototype = proto = new Prototype;
		aPrototype->SetBase(aBase ? aBase : Object::sPrototype);
		aPrototype->mFlags |= Object::ClassPrototype | Object::NativeClassPrototype;
		aPrototype->SetOwnProp(_T("__Class"), ExprTokenType(aClassName));
		aPrototype->mSize = aClassSize;

		TCHAR* name = full_name + _stprintf(full_name, _T("%s.Prototype."), aClassName);
		for (int i = 0; i < aMemberCount; ++i)
		{
			const auto& member = aMember[i];
			_tcscpy(name, member.name);
			if (member.invokeType == IT_CALL)
			{
				if (auto func = Method_New(full_name, (ObjectMember)member, aPrototype))
					func->Release();
			}
			else
			{
				auto prop = aPrototype->DefineProperty(name);
				prop->MinParams = member.minParams;
				prop->MaxParams = member.maxParams;

				auto op_name = _tcschr(name, '\0');

				_tcscpy(op_name, _T(".Get"));
				auto tmp = member;
				tmp.invokeType = IT_GET;
				auto func = (BuiltInMethod*)Method_New(full_name, tmp, nullptr);
				if (func)
					func->mClass = aPrototype, prop->SetGetter(func), func->Release();

				if (member.invokeType == IT_SET)
				{
					_tcscpy(op_name, _T(".Set"));
					if (func = (BuiltInMethod*)Method_New(full_name, (ObjectMember)member, nullptr))
						func->mClass = aPrototype, prop->SetSetter(func), func->Release();
				}
			}
		}
	}

	auto class_obj = Object::CreateClass(aPrototype);
	class_obj->SetBase(Object::sClass);
	if (proto)
		proto->Release();
	_stprintf(full_name, _T("%s.Call"), aClassName);
	auto ctor = (BuiltInFunc*)Func_New(FuncEntry{ full_name,obj_ctor,1,MAX_FUNCTION_PARAMS,0,{0} });
	ctor->mData = aPrototype;
	class_obj->DefineMethod(_T("Call"), ctor);
	ctor->Release();
	return class_obj;
}

IObject* STDMETHODCALLTYPE IAhkApi::GetEnumerator(IObject* aObj, int aVarCount) {
	IObject* enumerator;
	if (::GetEnumerator(enumerator, ExprTokenType(aObj), aVarCount, false) == OK)
		return enumerator;
	return nullptr;
}

bool STDMETHODCALLTYPE IAhkApi::CallEnumerator(IObject* aEnumerator, ExprTokenType* aParam[], int aParamCount) {
	return ::CallEnumerator(aEnumerator, aParam, aParamCount, false) == CONDITION_TRUE;
}

IObject* STDMETHODCALLTYPE IAhkApi::Object_New(ObjectType aType, ExprTokenType* aParam[], int aParamCount) {
	if (!Object::sPrototype || aType >= ObjectType::MAXINDEX || aType < ObjectType::Object)
		return nullptr;
	Object* obj = sObject[(int)aType];
	if (!obj) {
		if (aType == ObjectType::VarRef)
			return new VarRef;
		else if (aType == ObjectType::Module)
			return (IObject*)Module::Create(aParamCount ? TokenToString(*aParam[0]) : nullptr);
		else return nullptr;
	}
	ResultToken result{};
	TCHAR buf[MAX_NUMBER_SIZE];
	result.InitResult(buf);
	obj->Invoke(result, IT_CALL, nullptr, ExprTokenType(obj), aParam, aParamCount);
	if (result.symbol == SYM_OBJECT)
		return result.object;
	if (result.mem_to_free)
		free(result.mem_to_free);
	return nullptr;
}

bool STDMETHODCALLTYPE IAhkApi::Object_CallProp(ResultToken& aResultToken, Object* aObject, LPTSTR aName, ExprTokenType* aParam[], int aParamCount) {
	return aObject->Invoke(aResultToken, IT_CALL, aName, ExprTokenType(aObject), aParam, aParamCount) == OK;
}

bool STDMETHODCALLTYPE IAhkApi::Object_GetProp(ResultToken& aResultToken, Object* aObject, LPTSTR aName, bool aOwnProp, ExprTokenType* aParam[], int aParamCount) {
	if (aOwnProp && !aParamCount)
		if (aObject->GetOwnProp(aResultToken, aName)) {
			if (aResultToken.symbol == SYM_OBJECT)
				aResultToken.object->AddRef();
			return true;
		}
		else return false;
	return aObject->Invoke(aResultToken, IT_GET, aName, ExprTokenType(aObject), aParam, aParamCount) == OK;
}

bool STDMETHODCALLTYPE IAhkApi::Object_SetProp(Object* aObject, LPTSTR aName, ExprTokenType& aValue, bool aOwnProp, ExprTokenType* aParam[], int aParamCount) {
	if (aOwnProp && !aParamCount)
		return aObject->SetOwnProp(aName, aValue);
	ResultToken result{};
	result.SetValue(_T(""), 0);
	ExprTokenType** param = (ExprTokenType**)_alloca((aParamCount + 1) * sizeof(ExprTokenType*));
	param[0] = &aValue;
	if (aParamCount)
		memcpy(param + 1, aParam, aParamCount);
	bool ret = aObject->Invoke(result, IT_SET, aName, ExprTokenType(aObject), param, aParamCount + 1) == OK;
	result.Free();
	return ret;
}

bool STDMETHODCALLTYPE IAhkApi::Object_DeleteOwnProp(ResultToken& aResultToken, Object* aObject, LPTSTR aName) {
	ExprTokenType param(aName), * params = &param;
	aResultToken.symbol = SYM_MISSING;
	aObject->DeleteProp(aResultToken, 0, IT_CALL, &params, 1);
	return aResultToken.symbol != SYM_MISSING;
}

void STDMETHODCALLTYPE IAhkApi::Object_Clear(Object* aObject) { Object::FreesPrototype(aObject); }

bool STDMETHODCALLTYPE IAhkApi::Array_GetItem(Array* aArray, UINT aIndex, ExprTokenType& aValue) { return aArray->ItemToToken(aIndex, aValue); }

bool STDMETHODCALLTYPE IAhkApi::Array_SetItem(Array* aArray, UINT aIndex, ExprTokenType& aValue) {
	if (aIndex >= aArray->mLength)
		return Array_InsertItem(aArray, aValue, &aIndex);
	auto& field = aArray->mItem[aIndex];
	return field.Assign(aValue);
}

bool STDMETHODCALLTYPE IAhkApi::Array_InsertItem(Array* aArray, ExprTokenType& aValue, UINT* aIndex) {
	ExprTokenType* param[1] = { &aValue };
	return aIndex ? aArray->InsertAt(*aIndex, param, 1) == OK : aArray->Append(aValue);
}

bool STDMETHODCALLTYPE IAhkApi::Array_DeleteItem(ResultToken& aResultToken, Array* aArray, UINT aIndex) {
	if (aIndex >= aArray->mLength)
		return false;
	auto& item = aArray->mItem[aIndex];
	item.ReturnMove(aResultToken);
	item.AssignMissing();
	return true;
}

bool STDMETHODCALLTYPE IAhkApi::Array_RemoveItems(Array* aArray, UINT aIndex, UINT aCount) {
	if (aIndex >= aArray->mLength)
		return false;
	aArray->RemoveAt(aIndex, min(aCount, aArray->mLength - aIndex));
	return true;
}

Array* STDMETHODCALLTYPE IAhkApi::Array_FromEnumerable(IObject* aEnumerable, UINT aIndex) {
	IObject* enumerator;
	auto result = ::GetEnumerator(enumerator, ExprTokenType(aEnumerable), ++aIndex, false);
	if (result == FAIL || result == EARLY_EXIT)
		return nullptr;
	auto varref = new VarRef();
	ExprTokenType* param = (ExprTokenType*)_alloca(aIndex * (sizeof(ExprTokenType) + sizeof(ExprTokenType*)));
	ExprTokenType** params = (ExprTokenType**)(param + aIndex * sizeof(ExprTokenType));
	for (UINT i = 0; i < aIndex; i++)
		param[i].symbol = SYM_MISSING, params[i] = &param[i];
	param[aIndex - 1].SetValue(varref);
	Array* vargs = Array::Create();
	for (;;)
	{
		auto result = ::CallEnumerator(enumerator, &param, aIndex, false);
		if (result == FAIL)
		{
			vargs->Release();
			vargs = nullptr;
			break;
		}
		if (result != CONDITION_TRUE)
			break;
		ExprTokenType value;
		varref->ToTokenSkipAddRef(value);
		vargs->Append(value);
	}
	varref->Release();
	enumerator->Release();
	return vargs;
}

bool STDMETHODCALLTYPE IAhkApi::Buffer_Resize(BufferObject* aBuffer, size_t aSize) {
	return aBuffer->Resize(aSize) == OK;
}

bool STDMETHODCALLTYPE IAhkApi::Map_GetItem(Map* aMap, ExprTokenType& aKey, ExprTokenType& aValue) { return aMap->GetItem(aValue, aKey); }

bool STDMETHODCALLTYPE IAhkApi::Map_SetItem(Map* aMap, ExprTokenType& aKey, ExprTokenType& aValue) { return aMap->SetItem(aKey, aValue); }

bool STDMETHODCALLTYPE IAhkApi::Map_DeleteItem(ResultToken& aResultToken, Map* aMap, ExprTokenType& aKey) {
	ExprTokenType* param = &aKey;
	if (!Map_GetItem(aMap, aKey, ExprTokenType()))
		return false;
	aMap->Delete(aResultToken, 0, IT_CALL, &param, 1);
	return true;
}

void STDMETHODCALLTYPE IAhkApi::Map_Clear(Map* aMap) { aMap->Clear(); }

Object* STDMETHODCALLTYPE IAhkApi::JSON_Parse(LPTSTR aJSON) {
	JSON js;
	ExprTokenType param(aJSON), * params[] = { &param };
	ResultToken result{};
	js.Parse(result, params, 1);
	return result.symbol == SYM_OBJECT ? (Object*)result.object : nullptr;
}

LPTSTR STDMETHODCALLTYPE IAhkApi::JSON_Stringify(IObject* aObject, LPTSTR aIndent) {
	JSON js;
	ExprTokenType t_this(aObject), param(aIndent), * params[] = { &t_this, &param };
	ResultToken result{};
	js.Stringify(result, params, aIndent ? 2 : 1);
	return result.mem_to_free ? result.marker : nullptr;
}

void STDMETHODCALLTYPE IAhkApi::PumpMessages() {
	if (g_script && (sInit & 2)) {
		if (!g_script->mExiting)
			MsgSleep(SLEEP_INTERVAL, WAIT_FOR_MESSAGES);
		g_script->TerminateApp(EXIT_EXIT, g_ExitCode = 0);
	}
}


EXPORT IAhkApi* ahkGetApi(void* options) {
	if (options) {
		// {3066041C-D884-412A-A304-D2A452DD6B3A}
		IID_IObjectComCompatible = { 0x3066041c, 0xd884, 0x412a, { 0xa3, 0x4, 0xd2, 0xa4, 0x52, 0xdd, 0x6b, 0x3a } };
	}
	return IAhkApi::Initialize(); 
}
