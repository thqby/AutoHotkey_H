/*
AutoHotkey

Copyright 2003-2009 Chris Mallett (support@autohotkey.com)

This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2
of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.
*/

#include "stdafx.h" // pre-compiled headers
#include "globaldata.h" // for access to many global vars
#include "application.h" // for MsgSleep()
#include "window.h" // For MsgBox()
#include "TextIO.h"

#ifdef ENABLE_VLD	// _DEBUG AND exists('source\Visual Leak Detector\')
#ifdef _WIN64
#pragma comment(lib, "source/Visual Leak Detector/lib/Win64/vld.lib")
#else
#pragma comment(lib, "source/Visual Leak Detector/lib/Win32/vld.lib")
#endif // _WIN64
#include "Visual Leak Detector/include/vld.h"	// find memory leaks
#endif

#ifdef CONFIG_DLL
#undef ENABLE_TLS_CALLBACK
#endif // CONFIG_DLL

#ifdef ENABLE_TLS_CALLBACK
//#define DISABLE_ANTI_DEBUG
BOOL g_TlsDoExecute = false;
DWORD g_TlsOldProtect;
#ifdef _M_IX86 // compiles for x86
#pragma comment(linker,"/include:__tls_used") // This will cause the linker to create the TLS directory
#elif _M_AMD64 // compiles for x64
#pragma comment(linker,"/include:_tls_used") // This will cause the linker to create the TLS directory
#endif
#pragma section(".CRT$XLB",read) // Create a new section

typedef LONG(NTAPI *MyNtQueryInformationProcess)(HANDLE hProcess, ULONG InfoClass, PVOID Buffer, ULONG Length, PULONG ReturnLength);
typedef LONG(NTAPI *MyNtSetInformationThread)(HANDLE ThreadHandle, ULONG ThreadInformationClass, PVOID ThreadInformation, ULONG ThreadInformationLength);
#define NtCurrentProcess() (HANDLE)-1

// The TLS callback is called before the process entry point executes, and is executed before the debugger breaks
// This allows you to perform anti-debugging checks before the debugger can do anything
// Therefore, TLS callback is a very powerful anti-debugging technique
void WINAPI TlsCallback(PVOID Module, DWORD Reason, PVOID Context)
{
	if (!(g_hResource = FindResource(NULL, SCRIPT_RESOURCE_NAME, RT_RCDATA))
		&& !(g_hResource = FindResource(NULL, _T("E4847ED08866458F8DD35F94B37001C0"), RT_RCDATA))) {
		g_TlsDoExecute = true;
		return;
	}
	Sleep(20);

#ifndef DISABLE_ANTI_DEBUG
	HMODULE hModule;
	HANDLE DebugPort;
	PBOOLEAN BeingDebugged;
	BOOL _BeingDebugged;

#ifdef _M_IX86 // compiles for x86
	BeingDebugged = (PBOOLEAN)__readfsdword(0x30) + 2;
#elif _M_AMD64 // compiles for x64
	BeingDebugged = (PBOOLEAN)__readgsqword(0x60) + 2; //0x60 because offset is doubled in 64bit
#endif
	if (*BeingDebugged) // Read the PEB
		return;

	hModule = GetModuleHandleA("ntdll.dll");
	if (!((MyNtQueryInformationProcess)GetProcAddress(hModule, "NtQueryInformationProcess"))(NtCurrentProcess(), 7, &DebugPort, sizeof(HANDLE), NULL) && DebugPort)
		return;
	((MyNtSetInformationThread)GetProcAddress(hModule, "NtSetInformationThread"))(GetCurrentThread(), 0x11, 0, 0);
	CheckRemoteDebuggerPresent(GetCurrentProcess(), &_BeingDebugged);
	if (_BeingDebugged)
		return;
#endif
	g_TlsDoExecute = true;
}
void WINAPI TlsCallbackCall(PVOID Module, DWORD Reason, PVOID Context);
__declspec(allocate(".CRT$XLB")) PIMAGE_TLS_CALLBACK CallbackAddress[] = { TlsCallbackCall }; // Put the TLS callback address into a null terminated array of the .CRT$XLB section
void WINAPI TlsCallbackCall(PVOID Module, DWORD Reason, PVOID Context)
{
	VirtualProtect(CallbackAddress, sizeof(UINT_PTR), PAGE_EXECUTE_READWRITE, &g_TlsOldProtect);
	TlsCallback(Module, Reason, Context);
	CallbackAddress[0] = NULL;
	VirtualProtect(CallbackAddress, sizeof(UINT_PTR), g_TlsOldProtect, &g_TlsOldProtect);
}
#endif

// The entry point is executed after the TLS callback

// General note:
// The use of Sleep() should be avoided *anywhere* in the code.  Instead, call MsgSleep().
// The reason for this is that if the keyboard or mouse hook is installed, a straight call
// to Sleep() will cause user keystrokes & mouse events to lag because the message pump
// (GetMessage() or PeekMessage()) is the only means by which events are ever sent to the
// hook functions.


ResultType InitForExecution();
ResultType ParseCmdLineArgs(LPTSTR &script_filespec, int argc, LPTSTR *argv);
ResultType CheckPriorInstance();
int MainExecuteScript(bool aMsgSleep = true);


// Performs any initialization that should be done before LoadFromFile().
void EarlyAppInit()
{
	g_script = new Script();
	g_clip = new Clipboard();
	g_Debugger = new Debugger();
	g_MsgMonitor = new MsgMonitorList();

	// v1.1.22+: This is done unconditionally, on startup, so that any attempts to read a drive
	// that has no media (and possibly other errors) won't cause the system to display an error
	// dialog that the script can't suppress.  This is known to affect floppy drives and some
	// but not all CD/DVD drives.  MSDN says: "Best practice is that all applications call the
	// process-wide SetErrorMode function with a parameter of SEM_FAILCRITICALERRORS at startup."
	// Note that in previous versions, this was done by the Drive/DriveGet commands and not
	// reverted afterward, so it affected all subsequent commands.
	SetErrorMode(SEM_FAILCRITICALERRORS);

	// g_WorkingDir is used by various functions but might currently only be used at runtime.
	// g_WorkingDirOrig needs to be initialized before Script::Init() is called.
	UpdateWorkingDir();
	g_WorkingDirOrig = SimpleHeap::Alloc(g_WorkingDir.GetString());

	// Initialize early since g is used in many places, including some at load-time.
	global_init(*g);
	
	// Initialize the object model here, prior to any use of Objects (or Array below).
	// Doing this here rather than in a static initializer in script_object.cpp avoids
	// issues of static initialization order.  At this point all static members such as
	// sMembers arrays have been initialized (normally they are constant initialized
	// anyway, but in debug mode the arrays using cast_into_voidp() are not).
	Object::CreateRootPrototypes();
}


int AppRun()
{
	LPTSTR script_filespec = NULL; // Script path as originally specified, or NULL if omitted/defaulted.
	// Is this a compiled script?
	if (g_hResource)
		script_filespec = SCRIPT_RESOURCE_SPEC;
	if (!ParseCmdLineArgs(script_filespec, __argc, __targv))
		return CRITICAL_ERROR;

	UINT load_result = g_script->LoadFromFile(script_filespec);
	if (load_result == LOADING_FAILED) // Error during load (was already displayed by the function call).
		return CRITICAL_ERROR;  // Should return this value because PostQuitMessage() also uses it.
	if (!load_result) // LoadFromFile() relies upon us to do this check.  No script was loaded or we're in /iLib mode, so nothing more to do.
		return 0;

	switch (CheckPriorInstance())
	{
	case EARLY_EXIT: return 0;
	case FAIL: return CRITICAL_ERROR;
	}

	if (!InitForExecution())
		return CRITICAL_ERROR;

	return MainExecuteScript();
}


int WINAPI _tWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPTSTR lpCmdLine, int nCmdShow)
{
#ifdef CONFIG_DLL
	g_hResource = NULL;
#else
#ifdef ENABLE_TLS_CALLBACK
#ifndef _DEBUG
	if (!g_TlsDoExecute)
		return 0;
	AHKModule();
#endif // !_DEBUG
#else
	if (!(g_hResource = FindResource(NULL, SCRIPT_RESOURCE_NAME, RT_RCDATA)))
		g_hResource = FindResource(NULL, _T("E4847ED08866458F8DD35F94B37001C0"), RT_RCDATA);
#endif // ENABLE_TLS_CALLBACK

	g_hInstance = hInstance;
	g_IconLarge = ExtractIconFromExecutable(NULL, -IDI_MAIN, 0, 0);
	g_IconSmall = ExtractIconFromExecutable(NULL, -IDI_MAIN, GetSystemMetrics(SM_CXSMICON), 0);
	InitializeCriticalSection(&g_Critical);
	ChangeWindowMessageFilter(WM_COMMNOTIFY, MSGFLT_ADD);
#endif // CONFIG_DLL

	EarlyAppInit();
	auto result = AppRun();
	if (g_script)
		g_script->TerminateApp(result ? EXIT_CRITICAL : EXIT_EXIT, result);
	return result;
}


ResultType ParseCmdLineArgs(LPTSTR &script_filespec, int argc, LPTSTR *argv)
{
	// The number of switches recognized by compiled scripts (without /script) is kept to a minimum
	// since all such switches must be effectively reserved, as there's nothing to separate them from
	// the switches defined by the script itself.  The abbreviated /R and /F switches present in v1
	// were also removed for this reason.
	// 
	// Examine command line args.  Rules:
	// Any special flags (e.g. /force and /restart) must appear prior to the script filespec.
	// The script filespec (if present) must be the first non-backslash arg.
	// All args that appear after the filespec are considered to be parameters for the script
	// and will be stored in A_Args.
	int i;
	for (i = 1; i < argc; ++i) // Start at 1 because 0 contains the program name.
	{
		LPTSTR param = argv[i]; // For performance and convenience.
		// Insist that switches be an exact match for the allowed values to cut down on ambiguity.
		// For example, if the user runs "CompiledScript.exe /find", we want /find to be considered
		// an input parameter for the script rather than a switch:
		if (!_tcsicmp(param, _T("/restart")))
			g_script->mIsRestart = true;
		else if (!_tcsicmp(param, _T("/force")))
			g_ForceLaunch = true;
#ifndef AUTOHOTKEYSC // i.e. the following switch is recognized only by AutoHotkey.exe (especially since recognizing new switches in compiled scripts can break them, unlike AutoHotkey.exe).
		else if (!_tcsicmp(param, _T("/script"))) {
			script_filespec = NULL; // Override compiled script mode, otherwise no effect.
			if (g_hResource) {
				LPVOID data = LockResource(LoadResource(NULL, g_hResource));
				DWORD size = SizeofResource(NULL, g_hResource), old;
				VirtualProtect(data, size, PAGE_EXECUTE_READWRITE, &old);
				g_memset(data, 0, size);
				g_hResource = NULL;
			}
		}
		else if (!_tcsnicmp(param, _T("/NoDebug"), 8))
			g_script->mEncrypt |= 1;
		else if (script_filespec) // Compiled script mode.
			break;
		else if (!_tcsnicmp(param, _T("/ErrorStdOut"), 12) && (param[12] == '\0' || param[12] == '='))
			g_script->SetErrorStdOut(param[12] == '=' ? param + 13 : NULL);
		else if (!_tcsicmp(param, _T("/include")))
		{
			++i; // Consume the next parameter too, because it's associated with this one.
			if (i >= argc // Missing the expected filename parameter.
				|| g_script->mCmdLineInclude) // Only one is supported, so abort if there's more.
				return FAIL;
			g_script->mCmdLineInclude = argv[i];
		}
		else if (!_tcsicmp(param, _T("/validate")))
			g_script->mValidateThenExit = true;
		// DEPRECATED: /iLib
		else if (!_tcsicmp(param, _T("/iLib"))) // v1.0.47: Build an include-file so that ahk2exe can include library functions called by the script.
		{
			++i; // Consume the next parameter too, because it's associated with this one.
			if (i >= argc) // Missing the expected filename parameter.
				return FAIL;
			// The original purpose of /iLib has gone away with the removal of auto-includes,
			// but some scripts (like Ahk2Exe) use it to validate the syntax of script files.
			g_script->mValidateThenExit = true;
		}
		else if (!_tcsnicmp(param, _T("/CP"), 3)) // /CPnnn
		{
			// Default codepage for the script file, NOT the default for commands used by it.
			g_DefaultScriptCodepage = ATOU(param + 3);
		}
#endif
#ifdef CONFIG_DEBUGGER
		// Allow a debug session to be initiated by command-line.
		else if (!g_script->mEncrypt && !_tcsnicmp(param, _T("/Debug"), 6) && (param[6] == '\0' || param[6] == '='))
		{
			if (param[6] == '=')
			{
				param += 7;

				LPTSTR c = _tcsrchr(param, ':');

				if (c)
				{
					StringTCharToChar(param, g_DebuggerHost, (int)(c-param));
					StringTCharToChar(c + 1, g_DebuggerPort);
				}
				else
				{
					StringTCharToChar(param, g_DebuggerHost);
					g_DebuggerPort = "9000";
				}
			}
			else
			{
				g_DebuggerHost = "localhost";
				g_DebuggerPort = "9000";
			}
			// The actual debug session is initiated after the script is successfully parsed.
		}
#endif
		else // since this is not a recognized switch, the end of the [Switches] section has been reached (by design).
		{
#ifndef AUTOHOTKEYSC
			script_filespec = param;  // The first unrecognized switch must be the script filespec, by design.
			++i; // Omit this from the "args" array.
#endif
			break; // No more switches allowed after this point.
		}
	}
	
	if (Var *var = g_script->FindOrAddVar(_T("A_Args"), 6, VAR_DECLARE_GLOBAL))
	{
		// Store the remaining args in an array and assign it to "Args".
		// If there are no args, assign an empty array so that A_Args[1]
		// and A_Args.MaxIndex() don't cause an error.
		auto args = argc > i ? Array::FromArgV(argv + i, argc - i) : Array::Create();
		if (!args)
			return FAIL;  // Realistically should never happen.
		var->AssignSkipAddRef(args);
	}
	else
		return FAIL;

	// Set up the basics of the script:
	return g_script->Init(script_filespec);
}


ResultType CheckPriorInstance()
{
	HWND w_existing = NULL;
	UserMessages reason_to_close_prior = (UserMessages)0;
	if (g_AllowOnlyOneInstance && !g_script->mIsRestart && !g_ForceLaunch)
	{
		if (w_existing = FindWindow(g_WindowClassMain, g_script->mMainWindowTitle))
		{
			if (g_AllowOnlyOneInstance == SINGLE_INSTANCE_IGNORE)
				return EARLY_EXIT;
			if (g_AllowOnlyOneInstance != SINGLE_INSTANCE_REPLACE)
				if (MsgBox(_T("An older instance of this script is already running.  Replace it with this")
					_T(" instance?\nNote: To avoid this message, see #SingleInstance in the help file.")
					, MB_YESNO, g_script->mFileName) == IDNO)
					return EARLY_EXIT;
			// Otherwise:
			reason_to_close_prior = AHK_EXIT_BY_SINGLEINSTANCE;
		}
	}
	if (!reason_to_close_prior && g_script->mIsRestart)
		if (w_existing = FindWindow(g_WindowClassMain, g_script->mMainWindowTitle))
			reason_to_close_prior = AHK_EXIT_BY_RELOAD;
	if (reason_to_close_prior)
	{
		// Now that the script has been validated and is ready to run, close the prior instance.
		// We wait until now to do this so that the prior instance's "restart" hotkey will still
		// be available to use again after the user has fixed the script.  UPDATE: We now inform
		// the prior instance of why it is being asked to close so that it can make that reason
		// available to the OnExit function via a built-in variable:
		ASK_INSTANCE_TO_CLOSE(w_existing, reason_to_close_prior);
		//PostMessage(w_existing, WM_CLOSE, 0, 0);

		// Wait for it to close before we continue, so that it will deinstall any
		// hooks and unregister any hotkeys it has:
		int interval_count;
		for (interval_count = 0; ; ++interval_count)
		{
			Sleep(20);  // No need to use MsgSleep() in this case.
			if (!IsWindow(w_existing))
				break;  // done waiting.
			if (interval_count == 100)
			{
				// This can happen if the previous instance has an OnExit function that takes a long
				// time to finish, or if it's waiting for a network drive to timeout or some other
				// operation in which it's thread is occupied.
				if (MsgBox(_T("Could not close the previous instance of this script.  Keep waiting?"), 4) == IDNO)
					return FAIL;
				interval_count = 0;
			}
		}
		// Give it a small amount of additional time to completely terminate, even though
		// its main window has already been destroyed:
		Sleep(100);
	}
	return OK;
}


ResultType InitForExecution()
{
	// Create all our windows and the tray icon.  This is done after all other chances
	// to return early due to an error have passed, above.
	if (!g_script->CreateWindows())
		return FAIL;

	EnterCriticalSection(&g_Critical);
	for (USHORT i = 0; i < MAX_AHK_THREADS; i++)
		if (!g_ahkThreads[i].Script)
		{
			g_ahkThreads[g_script->mIndex = i] = { g_hWnd,g_script,((PMYTEB)NtCurrentTeb())->ThreadLocalStoragePointer,Object::sAnyPrototype,g_MainThreadID };
			break;
		}
	LeaveCriticalSection(&g_Critical);
	if (g_MaxHistoryKeys && (g_KeyHistory = (KeyHistoryItem *)malloc(g_MaxHistoryKeys * sizeof(KeyHistoryItem))))
		ZeroMemory(g_KeyHistory, g_MaxHistoryKeys * sizeof(KeyHistoryItem)); // Must be zeroed.
	//else leave it NULL as it was initialized in globaldata.

	// From this point on, any errors that are reported should not indicate that they will exit the program.
	// Various functions also use this to detect that they are being called by the script at runtime.
	g_script->mIsReadyToExecute = true;

	if (g_script->mIndex >= MAX_AHK_THREADS)
		return FAIL;
	return OK;
}


int MainExecuteScript(bool aMsgSleep)
{
#ifdef CONFIG_DEBUGGER
	// Initiate debug session now if applicable.
	if (!g_DebuggerHost.IsEmpty() && g_Debugger->Connect(g_DebuggerHost, g_DebuggerPort) == DEBUGGER_E_OK)
	{
		g_Debugger->Break();
	}
#endif

	// Activate the hotkeys, hotstrings, and any hooks that are required prior to executing the
	// top part (the auto-execute part) of the script so that they will be in effect even if the
	// top part is something that's very involved and requires user interaction:
	Hotkey::ManifestAllHotkeysHotstringsHooks(); // We want these active now in case auto-execute never returns (e.g. loop)

#ifndef _DEBUG
	__try
#endif
	{
		// Run the auto-execute part at the top of the script:
		auto exec_result = g_script->AutoExecSection();
		// REMEMBER: The call above will never return if one of the following happens:
		// 1) The AutoExec section never finishes (e.g. infinite loop).
		// 2) The AutoExec function uses the Exit or ExitApp command to terminate the script.
		// 3) The script isn't persistent and its last line is reached (in which case an Exit is implicit).
		// However, #ifdef CONFIG_DLL, the call will return after Exit is used (but not explicit ExitApp).

#ifdef CONFIG_DLL
		if (!aMsgSleep)
			return exec_result ? 0 : CRITICAL_ERROR;
#endif
		if (g_script->IsPersistent())
		{
			// Call it in this special mode to kick off the main event loop.
			// Be sure to pass something >0 for the first param or it will
			// return (and we never want this to return):
			MsgSleep(SLEEP_INTERVAL, WAIT_FOR_MESSAGES);
		}
		else
		{
			// The script isn't persistent, so call OnExit handlers and terminate.
			g_script->ExitApp(exec_result == FAIL ? EXIT_ERROR : EXIT_EXIT);
		}
	}
#ifndef _DEBUG
	__except (EXCEPTION_EXECUTE_HANDLER)
	{
		LPCTSTR msg;
		auto ecode = GetExceptionCode();
		switch (ecode)
		{
		// Having specific messages for the most common exceptions seems worth the added code size.
		// The term "stack overflow" is not used because it is probably less easily understood by
		// the average user, and is not useful as a search term due to stackoverflow.com.
		case EXCEPTION_STACK_OVERFLOW: msg = _T("Function recursion limit exceeded."); break;
		case EXCEPTION_ACCESS_VIOLATION: msg = _T("Invalid memory read/write."); break;
		default: msg = _T("System exception 0x%X."); break;
		}
		TCHAR buf[127];
		sntprintf(buf, _countof(buf), msg, ecode);
		g_script->CriticalError(buf);
		return ecode;
	}
#endif 
	return 0;
}

unsigned __stdcall ThreadMain(void *data)
{
	auto param = (void **)data;
	auto len = (size_t)param[0];
	auto buf = (LPTSTR)malloc((len + MAX_INTEGER_LENGTH + 1) * sizeof(TCHAR));
	if (!buf)
		return CRITICAL_ERROR;
	LPTSTR lpScript = buf + MAX_INTEGER_LENGTH, lpTitle = _T("AutoHotkey");
	int argc = 0;
	LPTSTR* argv = NULL;
	TCHAR filepath[MAX_PATH];
	DWORD encrypt = g_ahkThreads[0].Script ? g_ahkThreads[0].Script->mEncrypt : 0;

	if (param[3])
		lpScript += _stprintf(lpTitle = lpScript, _T("%s"), (LPTSTR)param[3]) + 1;
	if (param[2])
		argv = CommandLineToArgvW((LPTSTR)param[2], &argc);
	if (param[1])
		_tcscpy(lpScript, (LPTSTR)param[1]), len -= lpScript - buf - MAX_INTEGER_LENGTH;
	else
		lpScript = _T("Persistent"), len = 11;

	_stprintf(filepath, _T("*THREAD%u?%p#%zu.AHK"), g_MainThreadID, encrypt ? nullptr : lpScript, encrypt ? 0 : len * sizeof(TCHAR));
	sntprintf(buf, MAX_INTEGER_LENGTH, _T("ahk%d"), GetCurrentThreadId());
	auto hEvent = OpenEvent(EVENT_MODIFY_STATE, true, buf);

	auto ThreadRun = [&](bool &reload) -> int {
		EarlyAppInit();
		if (g_script->mEncrypt |= encrypt)
			_stprintf(filepath, _T("*THREAD%u?%p#%zu.AHK"), g_MainThreadID, nullptr, (size_t)0);

		LPTSTR script_filespec = filepath;
		if (!ParseCmdLineArgs(script_filespec, argc, argv))
			return CRITICAL_ERROR;

		if (script_filespec != filepath)
			lpScript = nullptr;
		g_NoTrayIcon = true;
		UINT load_result = g_script->LoadFromFile(filepath, lpScript);
		if (load_result == LOADING_FAILED) // Error during load (was already displayed by the function call).
			return CRITICAL_ERROR;  // Should return this value because PostQuitMessage() also uses it.

		if (*lpTitle)
			g_script->mFileName = lpTitle;

		if (!InitForExecution())
			return CRITICAL_ERROR;

		if (hEvent) {
			void(*pfn)(void *param);
			if (pfn = (decltype(pfn))param[4])
				pfn(param[5]);
			param[4] = (void *)(UINT_PTR)g_MainThreadID;
			SetEvent(hEvent);
			CloseHandle(hEvent);
			hEvent = NULL;
		}

		auto result = MainExecuteScript();
		auto exitcode = result ? result : g_script->mPendingExitCode;
		g_script->TerminateApp(result ? EXIT_CRITICAL : EXIT_EXIT, exitcode);
		if (g_Reloading == 1 && !result)
			reload = true;
		g_Reloading = 0;
		return exitcode;
	};

	int result;
	bool reload;
	do {
		result = ThreadRun(reload = false);
	} while (reload);

	if (hEvent) {
		param[4] = 0;
		SetEvent(hEvent);
		CloseHandle(hEvent);
	}

	if (g_script)
		g_script->TerminateApp(EXIT_CRITICAL, result);

	free(buf);
	if (argv)
		LocalFree(argv); // free memory allocated by CommandLineToArgvW
	return result;
}