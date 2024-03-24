#include "stdafx.h" // pre-compiled headers
#include "globaldata.h" // for access to many global vars
#include "application.h" // for InitNewThread()
#include "script_com.h"
#include "exports.h"
#include <process.h>  // NewThread
#include "LiteZip.h"  // crc32"
#include "MdFunc.h"

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
	Line *aFirstLine = g_script->mFirstLine,*aLastLine = g_script->mLastLine,*aCurrLine = g_script->mCurrLine;\
	UserFunc *aCurrFunc  = g->CurrentFunc;\
	int aClassObjectCount = g_script->mClassObjectCount;\
	auto aHotkeyCount = Hotkey::sHotkeyCount;\
	auto aHotstringCount = Hotstring::sHotstringCount;\
	g_script->mClassObjectCount = 0;g_script->mFirstLine = g_script->mLastLine = NULL;g->CurrentFunc = NULL;

#define RESTORE_G_SCRIPT \
	g_script->mFirstLine = aFirstLine;\
	g_script->mLastLine = aLastLine;\
	g_script->mLastLine->mNextLine = NULL;\
	g_script->mCurrLine = aCurrLine;\
	g_script->mClassObjectCount = aClassObjectCount + g_script->mClassObjectCount;\
	g_script->mCurrFileIndex = aCurrFileIndex;\
	g_script->mCombinedLineNumber = aCombinedLineNumber;


void callFuncDll(FuncAndToken &aToken, bool throwerr)
{
	if (g_nThreads >= g_MaxThreadsTotal) {
		aToken.mResult->SetResult(FAIL);
		return;
	}
	IObject* func = aToken.mFunc;
	TCHAR buf[MAX_NUMBER_SIZE];
	aToken.mResult->InitResult(buf);

	// See MsgSleep() for comments about the following section.
	InitNewThread(0, false, true);
	g->ExcptMode = EXCPTMODE_CATCH;
	func->Invoke(*aToken.mResult, IT_CALL, nullptr, ExprTokenType(func), aToken.mParam, aToken.mParamCount);
	if (g->ThrownToken) {
		if (throwerr)
			g_script->UnhandledException(NULL);
		else {
			aToken.mResult->CopyValueFrom(*g->ThrownToken);
			aToken.mResult->mem_to_free = g->ThrownToken->mem_to_free;
			g->ThrownToken->symbol = SYM_INTEGER;
			g->ThrownToken->mem_to_free = nullptr;
		}
	}
	if (aToken.mResult->symbol == SYM_STRING && !aToken.mResult->mem_to_free)
		aToken.mResult->Malloc(aToken.mResult->marker, aToken.mResult->marker_length);
	ResumeUnderlyingThread();
	return;
}

EXPORT(int) ahkReady(DWORD aThreadID) // HotKeyIt check if dll is ready to execute
{
#pragma comment(linker,"/export:" __FUNCTION__"=" __FUNCDNAME__)
	AutoTLS atls;
	return (int)(atls.Enter(aThreadID) && g_script && !g_Reloading && g_script->mIsReadyToExecute);
}

EXPORT(int) ahkPause(int aNewState, DWORD aThreadID) //Change pause state of a running script
{
#pragma comment(linker,"/export:" __FUNCTION__"=" __FUNCDNAME__)
	AutoTLS atls;
	if (atls.Enter(aThreadID)) {
		auto &IsPaused = g->IsPaused;
		auto hwnd = g_hWnd;
		if (aNewState == -1)
			aNewState = !IsPaused;
		if (!!aNewState == IsPaused)
			return (int)IsPaused;
		if (IsPaused)
			--g_nPausedThreads;
		else
			++g_nPausedThreads;
		IsPaused = !IsPaused;
		g_script->UpdateTrayIcon();
		if (!IsPaused)
			atls.~AutoTLS(), SendMessage(hwnd, WM_COMMNOTIFY, 0, 0);
		return (int)IsPaused;
	}
	return 0;
}

EXPORT(UINT_PTR) ahkFindFunc(LPTSTR aFuncName, DWORD aThreadID) {
#pragma comment(linker,"/export:" __FUNCTION__"=" __FUNCDNAME__)
	AutoTLS atls;
	return atls.Enter(aThreadID) ? (UINT_PTR)g_script->FindGlobalFunc(aFuncName) : 0;
}

EXPORT(UINT_PTR) ahkFindLabel(LPTSTR aLabelName, DWORD aThreadID) {
#pragma comment(linker,"/export:" __FUNCTION__"=" __FUNCDNAME__)
	AutoTLS atls;
	if (atls.Enter(aThreadID))
		return (UINT_PTR)g_script->FindLabel(aLabelName);
	return 0;
}

// Naveen: v1. ahkGetVar()
EXPORT(LPTSTR) ahkGetVar(LPTSTR name, int getVar, DWORD aThreadID)
{
#pragma comment(linker,"/export:" __FUNCTION__"=" __FUNCDNAME__)
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
				ExprTokenType exp;
				ahkvar->ToTokenSkipAddRef(exp);
				auto str = TokenToString(exp, buf);
				return _tcsdup(str);
			}
		}
	}
	return NULL;
}

EXPORT(int) ahkAssign(LPTSTR name, LPTSTR value, DWORD aThreadID)
{
#pragma comment(linker,"/export:" __FUNCTION__"=" __FUNCDNAME__)
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
			DWORD_PTR res;
			atls.~AutoTLS();
			return SendMessageTimeout(hwnd, AHK_THREADVAR, (WPARAM)&token, (LPARAM)ahkvar,
				SMTO_ABORTIFHUNG, 5000, &res) && res;
		}
		return ahkvar->Assign(value) == OK;
	}
	return 0;
}

//HotKeyIt ahkExecuteLine()
EXPORT(UINT_PTR) ahkExecuteLine(UINT_PTR line, int aMode, int wait, DWORD aThreadID)
{
#pragma comment(linker,"/export:" __FUNCTION__"=" __FUNCDNAME__)
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

EXPORT(int) ahkLabel(LPTSTR aLabelName, int nowait, DWORD aThreadID) // 0 = wait = default
{
#pragma comment(linker,"/export:" __FUNCTION__"=" __FUNCDNAME__)
	AutoTLS atls;
	if (atls.Enter(aThreadID)) {
		HWND hwnd = g_hWnd;
		auto label = g_script->FindLabel(aLabelName);
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
					vars[i]->Free(VAR_ALWAYS_FREE | VAR_CLEAR_ALIASES | VAR_REQUIRE_INIT), vars[i++] = NULL;
				vars[j++] = var;
			}
			list.mCount = mCount;
		}
		free(mItem), mCount = -1;
	}
};

// HotKeyIt: addScript()
UINT_PTR _addScript(LPTSTR script, int waitexecute, DWORD aThreadID, int _catch)
{
	AutoTLS atls;
	int ret = atls.Enter(aThreadID);
	if (!ret)
		return 0;
	else if (ret == 2) {
		ExprTokenType params[4] = { (__int64)_addScript, _T("i=siuii") };
		ExprTokenType* param[] = { params, params + 1, params + 2, params + 3 };
		ResultToken result;
		IObject* func = DynaToken::Create(param, 2);
		FuncAndToken token = { func,&result,param,4 };
		auto hwnd = g_hWnd;
		DWORD_PTR res;
		params[0].SetValue(script), params[1].SetValue(waitexecute), params[2].SetValue((UINT)aThreadID), params[3].SetValue(-_catch);
		result.InitResult(_T(""));
		atls.~AutoTLS();
		SendMessageTimeout(hwnd, AHK_EXECUTE_FUNCTION, (WPARAM)&token, 0, SMTO_ABORTIFHUNG, 5000, &res);
		func->Release();
		return (UINT_PTR)TokenToInt64(result);
	}
	EnterCriticalSection(&g_CriticalTLS);
	int aVarCount = g_script->mVars.mCount;
	int aFuncCount = g_script->mFuncs.mCount;
	VarListBackup aVarBkp(g_script->mVars);
	if (aVarBkp.mCount == -1) {
		LeaveCriticalSection(&g_CriticalTLS);
		return 0;
	}
	int HotkeyCount = Hotkey::sHotkeyCount;
	HotkeyCriterion* aFirstHotExpr = g_FirstHotExpr, * aLastHotExpr = g_LastHotExpr;
	g_FirstHotExpr = NULL; g_LastHotExpr = NULL;
	int aSourceFileIdx = Line::sSourceFileCount;
	
	// Backup SimpleHeap to restore later
	auto heapbkp = SimpleHeap::HeapBackUp();
	int excp = g->ExcptMode;
	TCHAR tmp[MAX_PATH];
	void* p = nullptr;
	if (_catch < 0)
		g->ExcptMode |= EXCPTMODE_CATCH, g->ExcptMode &= ~EXCPTMODE_CAUGHT;
	if (!g_script->mEncrypt)
		p = SimpleHeap::Alloc(script);
	_stprintf(tmp, _T("*THREAD%u?%p#%zu.AHK"), g_MainThreadID, p, _tcslen(script) * sizeof(TCHAR));
	BACKUP_G_SCRIPT
	if (g_script->LoadFromFile(tmp, script) != TRUE)
	{
		RESTORE_IF_EXPR
		g->ExcptMode = excp;
		g->CurrentFunc = (UserFunc*)aCurrFunc;
		Line* aTempLine = g_script->mLastLine;
		Line* aExecLine = g_script->mFirstLine;
		RESTORE_G_SCRIPT;
		if (g->ThrownToken && _catch < 0)
			Script::FreeExceptionToken(g->ThrownToken);
		Line::sSourceFileCount = aSourceFileIdx;
		aVarBkp.~VarListBackup();
		for (auto &i = Hotkey::sHotkeyCount; i > aHotkeyCount;)
			delete Hotkey::shk[--i];
		for (auto &i = Hotstring::sHotstringCount; i > aHotstringCount;)
			delete Hotstring::shs[--i];
		auto &fns = g_script->mFuncs;
		for (auto &i = fns.mCount; i > aFuncCount;) {
			auto &fn = *fns.mItem[--i];
			fn.mClass && fn.mClass->Release();
			fn.Release();
		}
		for (Line *line = aTempLine; line; line = line->mPrevLine)
			line->Free();
		Line::FreeDerefBufIfLarge();
		// restore SimpleHeap
		heapbkp.Restore();
		g_script->mIsReadyToExecute = true;
		LeaveCriticalSection(&g_CriticalTLS);
		return 0;
	}
	FINALIZE_HOTKEYS
	RESTORE_IF_EXPR
	aVarBkp.mCount = aVarBkp.list.mCount;
	g->CurrentFunc = (UserFunc*)aCurrFunc;
	Line *aTempLine = g_script->mFirstLine;
	aLastLine->mNextLine = aTempLine;
	aTempLine->mPrevLine = aLastLine;
	aLastLine = g_script->mLastLine;
	RESTORE_G_SCRIPT;
	g_script->mIsReadyToExecute = true;
	if (waitexecute) {
		if (waitexecute == 1) {
			aTempLine->ExecUntil(UNTIL_RETURN);
			if (g->ThrownToken && _catch < 0)
				Script::FreeExceptionToken(g->ThrownToken);
		}
		else
			PostMessage(g_hWnd, AHK_EXECUTE, (WPARAM)aTempLine, 1);
	}
	g->ExcptMode = excp;
	LeaveCriticalSection(&g_CriticalTLS);
	return (UINT_PTR)aTempLine;
}

EXPORT(UINT_PTR) addScript(LPTSTR script, int waitexecute, DWORD aThreadID) {
#pragma comment(linker,"/export:" __FUNCTION__"=" __FUNCDNAME__)
	return _addScript(script, waitexecute, aThreadID, int(g && (g->ExcptMode & EXCPTMODE_CATCH)));
}

int _ahkExec(LPTSTR script, DWORD aThreadID, int _catch)
{
	AutoTLS atls;
	int ret = atls.Enter(aThreadID);
	if (!ret)
		return 0;
	else if (ret == 2) {
		ExprTokenType params[3] = { (__int64)_ahkExec, _T("i=suii") };
		ExprTokenType *param[] = { params,params + 1,params + 2 };
		ResultToken result;
		IObject* func = DynaToken::Create(param, 2);
		FuncAndToken token = { func,&result,param,3 };
		auto hwnd = g_hWnd;
		DWORD_PTR res;
		params[0].SetValue(script), params[1].SetValue((__int64)aThreadID), params[2].SetValue(-_catch);
		result.InitResult(_T(""));
		atls.~AutoTLS();
		SendMessageTimeout(hwnd, AHK_EXECUTE_FUNCTION, (WPARAM)&token, 0, SMTO_ABORTIFHUNG, 5000, &res);
		func->Release();
		return (int)TokenToInt64(result);
	}
	EnterCriticalSection(&g_CriticalTLS);
	int aVarCount = g_script->mVars.mCount;
	int aFuncCount = g_script->mFuncs.mCount;
	VarListBackup aVarBkp(g_script->mVars);
	VarListBackup* fvbkp = nullptr, * fsvbkp = nullptr;
	if (aVarBkp.mCount == -1) {
		LeaveCriticalSection(&g_CriticalTLS);
		return 0;
	}
	if (!aThreadID) {
		if (auto cur = g->CurrentFunc) {
			fvbkp = new VarListBackup(cur->mVars);
			fsvbkp = new VarListBackup(cur->mStaticVars);
			if (fvbkp->mCount == -1 || fsvbkp->mCount == -1) {
				delete fvbkp;
				delete fsvbkp;
				LeaveCriticalSection(&g_CriticalTLS);
				return 0;
			}
		}
	}
	int HotkeyCount = Hotkey::sHotkeyCount;
	HotkeyCriterion *aFirstHotExpr = g_FirstHotExpr,*aLastHotExpr = g_LastHotExpr;
	g_FirstHotExpr = NULL;g_LastHotExpr = NULL;
	int aSourceFileIdx = Line::sSourceFileCount;
	BACKUP_G_SCRIPT

	// Backup SimpleHeap to restore later
	auto heapbkp = SimpleHeap::HeapBackUp(true);
	int excp = g->ExcptMode;
	if (_catch < 0)
		g->ExcptMode |= EXCPTMODE_CATCH, g->ExcptMode &= ~EXCPTMODE_CAUGHT;

	if (!aThreadID)
		g->CurrentFunc = aCurrFunc;

	TCHAR tmp[MAX_PATH];
	_stprintf(tmp, _T("*THREAD%u?%p#%zu.AHK"), g_MainThreadID, g_script->mEncrypt ? nullptr : script, _tcslen(script) * sizeof(TCHAR));
	auto result = g_script->LoadFromFile(tmp, script);
	FINALIZE_HOTKEYS
	RESTORE_IF_EXPR
	g->CurrentFunc = (UserFunc*)aCurrFunc;
	Line *aTempLine = g_script->mLastLine;
	Line *aExecLine = g_script->mFirstLine;
	RESTORE_G_SCRIPT;
	// restore SimpleHeap
	heapbkp.Restore();
	g_script->mIsReadyToExecute = true;
	
	if (result == TRUE) {
		g_script->mLastLine->mNextLine = aExecLine;
		aExecLine->ExecUntil(UNTIL_RETURN);
		g_script->mLastLine->mNextLine = NULL;
		g_script->mCurrLine = aCurrLine;
	}
	if (g->ThrownToken && _catch < 0)
		Script::FreeExceptionToken(g->ThrownToken);
	g->ExcptMode = excp;
	Line::sSourceFileCount = aSourceFileIdx;
	if (fvbkp) {
		delete fvbkp;
		delete fsvbkp;
	}
	aVarBkp.~VarListBackup();
	for (auto &i = Hotkey::sHotkeyCount; i > aHotkeyCount;)
		delete Hotkey::shk[--i];
	for (auto &i = Hotstring::sHotstringCount; i > aHotstringCount;)
		delete Hotstring::shs[--i];
	auto &fns = g_script->mFuncs;
	for (auto &i = fns.mCount; i > aFuncCount;) {
		auto &fn = *fns.mItem[--i];
		fn.mClass && fn.mClass->Release();
		fn.Release();
	}
	for (Line *line = aTempLine; line; line = line->mPrevLine)
		line->Free();
	Line::FreeDerefBufIfLarge();
	heapbkp.DeleteAfter(heapbkp.mFirst2);
	LeaveCriticalSection(&g_CriticalTLS);
	return result == TRUE;
}

EXPORT(int) ahkExec(LPTSTR script, DWORD aThreadID) {
#pragma comment(linker,"/export:" __FUNCTION__"=" __FUNCDNAME__)
	return _ahkExec(script, aThreadID, int(g && (g->ExcptMode & EXCPTMODE_CATCH)));
}

LPTSTR ahkFunction(LPTSTR func, LPTSTR param1, LPTSTR param2, LPTSTR param3, LPTSTR param4, LPTSTR param5, LPTSTR param6, LPTSTR param7, LPTSTR param8, LPTSTR param9, LPTSTR param10, DWORD aThreadID, int sendOrPost)
{
	AutoTLS atls;
	if (!atls.Enter(aThreadID))
		return NULL;
	Var* aVar = g_script->FindGlobalVar(func);
	Func* aFunc = aVar ? dynamic_cast<Func*>(aVar->ToObject()) : nullptr;
	if (!aFunc)
		return NULL;
	int aParamsCount = 10;
	LPTSTR *params[10] = { &param1,&param2,&param3,&param4,&param5,&param6,&param7,&param8,&param9,&param10 };
	for (; aParamsCount > 0; aParamsCount--)
		if (*params[aParamsCount - 1])
			break;
	if (aParamsCount < aFunc->mMinParams)
	{
		g_script->RuntimeError(ERR_TOO_FEW_PARAMS);
		return NULL;
	}
	Promise *promise = nullptr;
	ExprTokenType paramtoken[10];
	ExprTokenType *param[10] = { 0 };
	HWND hwnd = g_hWnd;
	aParamsCount = aFunc->mParamCount < aParamsCount && !aFunc->mIsVariadic ? aFunc->mParamCount : aParamsCount;

	for (int i = 0; i < aParamsCount; i++) {
		param[i] = &paramtoken[i];
		if (*params[i])
			paramtoken[i].SetValue(*params[i]);
		else paramtoken[i].symbol = SYM_MISSING;
	}
	atls.~AutoTLS();
	if (sendOrPost) {
		FuncResult result;
		FuncAndToken token = { aFunc, &result, param, aParamsCount };
		DWORD_PTR res;
		if (!hwnd || !atls.teb) {
			callFuncDll(token, true);
			if (result.symbol == SYM_OBJECT)
				result.object->Release();
		}
		else if (!SendMessageTimeout(hwnd, AHK_EXECUTE_FUNCTION, (WPARAM)&token, 0, SMTO_ABORTIFHUNG, 5000, &res))
			return NULL;
		if (result.Exited()) {
			free(result.mem_to_free);
			return NULL;
		}
		auto str = TokenToString(result, result.buf);
		if (!result.mem_to_free)
			str = _tcsdup(str);
		return str;
	}
	else {
		promise = new Promise(atls.teb ? nullptr : aFunc, aVar, false);
		if (!promise->Init(param, aParamsCount))
		{
			promise->Release();
			return 0;
		}
		if (g) promise->mPriority = g->Priority;
		bool ret;
		if (!(ret = PostMessage(hwnd, AHK_EXECUTE_FUNCTION, 0, (LPARAM)promise)))
			promise->Release();
		return (LPTSTR)ret;
	}
}

EXPORT(LPTSTR) ahkFunction(LPTSTR func, LPTSTR param1, LPTSTR param2, LPTSTR param3, LPTSTR param4, LPTSTR param5, LPTSTR param6, LPTSTR param7, LPTSTR param8, LPTSTR param9, LPTSTR param10, DWORD aThreadID)
{
#pragma comment(linker,"/export:" __FUNCTION__"=" __FUNCDNAME__)
	return ahkFunction(func, param1, param2, param3, param4, param5, param6, param7, param8, param9, param10, aThreadID, 1);
}

EXPORT(int) ahkPostFunction(LPTSTR func, LPTSTR param1, LPTSTR param2, LPTSTR param3, LPTSTR param4, LPTSTR param5, LPTSTR param6, LPTSTR param7, LPTSTR param8, LPTSTR param9, LPTSTR param10, DWORD aThreadID)
{
#pragma comment(linker,"/export:" __FUNCTION__"=" __FUNCDNAME__)
	return (int)(UINT_PTR)ahkFunction(func, param1, param2, param3, param4, param5, param6, param7, param8, param9, param10, aThreadID, 0);
}

unsigned __stdcall ThreadMain(void *data);
DWORD NewThread(LPTSTR aScript, LPTSTR aCmdLine, LPTSTR aTitle, BOOL aCatch, void(*callback)(void *param), void* param)
{
	void* data[7] = { 0,0,0,0,(void*)(UINT_PTR)aCatch,callback,param};
	size_t len = 0;
	if (aScript)
		data[1] = aScript, len += _tcsclen(aScript) + 1;
	if (aCmdLine && *aCmdLine)
		data[2] = aCmdLine;
	if (aTitle && *aTitle)
		data[3] = aTitle, len += _tcsclen(aTitle) + 1;
	data[0] = (void *)len;
	unsigned int ThreadID;
	TCHAR buf[MAX_INTEGER_LENGTH];
	HANDLE hThread = (HANDLE)_beginthreadex(NULL, 0, ThreadMain, data, 0, &ThreadID);
	if (hThread)
	{
		CloseHandle(hThread);
		sntprintf(buf, _countof(buf), _T("ahk%d"), ThreadID);
		if (HANDLE hEvent = CreateEvent(NULL, false, false, buf))
		{
			// we need to give the thread also little time to allocate memory to avoid std::bad_alloc that might happen when you run newthread in a loop
			WaitForSingleObject(hEvent, 2000);
			CloseHandle(hEvent);
		}
		else
			return ThreadID;
	}
	return data[5] ? ThreadID : 0;
}

EXPORT(DWORD) NewThread(LPTSTR aScript, LPTSTR aCmdLine, LPTSTR aTitle) // HotKeyIt check if dll is ready to execute
{
#pragma comment(linker,"/export:" __FUNCTION__"=" __FUNCDNAME__)
	return NewThread(aScript, aCmdLine, aTitle, FALSE, nullptr, nullptr);
}

IAhkApi IAhkApi::instance;
thread_local Object* IAhkApi::sObject[(int)ObjectType::MAXINDEX] = {};
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

void EarlyAppInit();
ResultType InitForExecution();
ResultType ParseCmdLineArgs(LPTSTR &script_filespec, int argc, LPTSTR *argv);
IAhkApi* IAhkApi::Initialize() {
	int fg = 0, i = 0;
#ifdef CONFIG_DLL
	if (!g_script) {
		LPTSTR script_file = _T("* ");
		EarlyAppInit();
		ParseCmdLineArgs(script_file, 0, nullptr);
		g_script->mFileName = g_WindowClassMain;
		g_script->LoadFromFile(script_file, _T(""));
		InitForExecution();
		fg = 2;
		g_script->AutoExecSection();
	}
#endif // CONFIG_DLL
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
		instance.sInit = 1 | fg;
#undef GETAHKOBJECT
	}
	return &instance;
}

void IAhkApi::Finalize() {
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

void STDMETHODCALLTYPE IAhkApi::VarToToken(Var* aVar, ExprTokenType& aToken) {
	if (aVar->mType == VAR_VIRTUAL) {
		FuncResult result_token;
		result_token.symbol = SYM_INTEGER; // For _f_return_i() and consistency with BIFs.
		aVar->Get(result_token);
		if (result_token.Exited())
			result_token.symbol = SYM_INVALID;
		else if (result_token.symbol == SYM_STRING) {
			if (result_token.mem_to_free)
				aVar->_AcceptNewMem(result_token.mem_to_free, result_token.marker_length);
			else {
				size_t length;
				LPTSTR value = TokenToString(result_token, result_token.buf, &length);
				if (aVar->AssignString(nullptr, length))
					tmemcpy(result_token.marker = aVar->mCharContents, value, (result_token.marker_length = length) + 1);
				else result_token.symbol = SYM_INVALID;
			}
			aVar->mAttrib &= ~VAR_ATTRIB_VIRTUAL_OPEN;
		}
		aToken.CopyValueFrom(result_token);
	}
	else
		aVar->ToTokenSkipAddRef(aToken);
}

void STDMETHODCALLTYPE IAhkApi::VarFree(Var* aVar, int aWhenToFree) { aVar->Free(aWhenToFree); }

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
		VarToToken(var, aValue);
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
	HMODULE mod = nullptr;
	size_t sz_name, sz_out;
	if (!GetModuleHandleEx(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, aBIF.mName, &mod))
		sz_name = (_tcslen(aBIF.mName) + 1) * sizeof(TCHAR);
	else sz_name = 0;
	if (aBIF.mOutputVars && !GetModuleHandleEx(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, (LPCTSTR)&aBIF, &mod))
		sz_out = MAX_FUNC_OUTPUT_VAR;
	else sz_out = 0;
	auto func = ::new(malloc(sizeof(BuiltInFunc) + sz_name + sz_out)) BuiltInFunc(aBIF);
	if (!func) return nullptr;
	func->MarkNeedFree();
	if (sz_out) {
		func->mOutputVars = (UCHAR *)func + sizeof(BuiltInFunc);
		memcpy(func->mOutputVars, aBIF.mOutputVars, MAX_FUNC_OUTPUT_VAR);
	}
	if (sz_name) {
		func->mName = (LPCTSTR)((char *)func + sizeof(BuiltInFunc) + sz_out);
		memcpy((void *)func->mName, aBIF.mName, sz_name);
	}
	return func;
}

Func* STDMETHODCALLTYPE IAhkApi::Method_New(LPTSTR aFullName, ObjectMember& aMember, Object* aPrototype) {
	HMODULE mod = nullptr;
	size_t sz_name;
	if (!GetModuleHandleEx(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, aFullName, &mod))
		sz_name = (_tcslen(aFullName) + 1) * sizeof(TCHAR);
	else sz_name = 0;
	auto func = ::new (malloc(sizeof(BuiltInMethod) + sz_name)) BuiltInMethod(aFullName);
	if (!func) return nullptr;
	func->MarkNeedFree();
	if (sz_name) {
		func->mName = (LPCTSTR)((char *)func + sizeof(BuiltInMethod));
		memcpy((void *)func->mName, aFullName, sz_name);
	}
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
			if (func->mMIT == IT_GET) {
				prop->SetGetter(func);
				prop->NoParamGet = aMember.maxParams == 0;
				prop->NoEnumGet = aMember.minParams > 0;
			}
			else {
				prop->SetSetter(func);
				prop->NoParamSet = aMember.maxParams == 0;
			}
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
				prop->NoParamGet = prop->NoParamSet = member.maxParams == 0;
				prop->NoEnumGet = member.minParams > 0;

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

class Module : public Object
{
	LPTSTR mName;
public:
	~Module() { if (mName) free(mName); }
	static Object *Create(LPTSTR aName = nullptr) {
		auto obj = new Module;
		obj->SetBase(Object::sPrototype);
		if (aName && *aName)
			obj->mName = _tcsdup(aName);
		else obj->mName = nullptr;
		return obj;
	}

	LPTSTR Type() { return mName ? mName : _T("Module"); }
	ResultType Invoke(IObject_Invoke_PARAMS_DECL) {
		if (aFlags & IT_CALL) {
			ResultToken this_token{};
			if (aName && GetOwnProp(this_token, aName) && this_token.symbol == SYM_OBJECT)
				return this_token.object->Invoke(aResultToken, IT_CALL, nullptr, this_token, aParam, aParamCount);
		}
		return Object::Invoke(IObject_Invoke_PARAMS);
	}
};

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
		if (!g_Reloading)
			MsgSleep(SLEEP_INTERVAL, WAIT_FOR_MESSAGES);
		g_script->TerminateApp(EXIT_EXIT, 0);
	}
}

Func *STDMETHODCALLTYPE IAhkApi::MdFunc_New(LPCTSTR aName, void *aFuncPtr, MdType *aSig, Object *aPrototype) {
	HMODULE mod = nullptr;
	size_t sz_name, sz_sig;
	UINT ac = 0;
	for (auto sig = aSig + 1; sig[ac] != MdType::Void; ++ac);
	if (!GetModuleHandleEx(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, aName, &mod))
		sz_name = (_tcslen(aName) + 1) * sizeof(TCHAR);
	else sz_name = 0;
	if (!GetModuleHandleEx(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, (LPCTSTR)aSig, &mod))
		sz_sig = ac;
	else sz_sig = 0;
	auto p = (char *)malloc(sizeof(MdFunc) + sz_name + sz_sig);
	if (!p) return nullptr;
	if (sz_name) {
		auto t = p + sizeof(MdFunc);
		memcpy(t, aName, sz_name);
		aName = (LPCTSTR)t;
	}
	auto arg = aSig + 1;
	if (sz_sig) {
		arg = (MdType*)(p + sizeof(MdFunc) + sz_name);
		memcpy(arg, aSig + 1, sz_sig);
	}
	auto func = ::new(p) MdFunc(aName, aFuncPtr, aSig[0], arg, ac, aPrototype);
	func->MarkNeedFree();
	return func;
}


EXPORT(IAhkApi*) ahkGetApi(void* options) {
#pragma comment(linker,"/export:" __FUNCTION__"=" __FUNCDNAME__)
	if (options) {
	}
	return IAhkApi::Initialize(); 
}
