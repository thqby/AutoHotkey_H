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

#include <iphlpapi.h>
#include <initializer_list>
#ifdef _DEBUG
#ifdef _WIN64
#pragma comment(lib, "D:/Program Files/Visual Leak Detector/lib/Win64/vld.lib")
#else
#pragma comment(lib, "D:/Program Files/Visual Leak Detector/lib/Win32/vld.lib")
#endif // _WIN64
#include <vld.h>	// find memory leaks
#endif

#ifdef ENABLE_TLS_CALLBACK
//#define DISABLE_ANTI_DEBUG
#ifndef _USRDLL
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
//#define DISABLE_ANTI_DEBUG
void WINAPI TlsCallback(PVOID Module, DWORD Reason, PVOID Context)
{
	int i = 0;
	for (auto p : { *(PULONGLONG)CryptHashData, *(PULONGLONG)CryptDeriveKey, *(PULONGLONG)CryptDestroyHash, *(PULONGLONG)CryptEncrypt, *(PULONGLONG)CryptDecrypt, *(PULONGLONG)CryptDestroyKey })
		g_crypt_code[i++] = p;

	if (!(g_hResource = FindResource(NULL, SCRIPT_RESOURCE_NAME, RT_RCDATA)))
		if (!(g_hResource = FindResource(NULL, _T("E4847ED08866458F8DD35F94B37001C0"), RT_RCDATA))) {
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
#endif
// The entry point is executed after the TLS callback

// General note:
// The use of Sleep() should be avoided *anywhere* in the code.  Instead, call MsgSleep().
// The reason for this is that if the keyboard or mouse hook is installed, a straight call
// to Sleep() will cause user keystrokes & mouse events to lag because the message pump
// (GetMessage() or PeekMessage()) is the only means by which events are ever sent to the
// hook functions.


int MainExecuteScript();
ResultType ThreadExecuteScript();

#ifndef _USRDLL
int WINAPI _tWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPTSTR lpCmdLine, int nCmdShow)
{
#ifdef ENABLE_TLS_CALLBACK
	if (!g_TlsDoExecute)
		return 0;
#else
	if (!g_crypt_code[0]) {
		int i = 0;
		for (auto p : { *(PULONGLONG)CryptHashData, *(PULONGLONG)CryptDeriveKey, *(PULONGLONG)CryptDestroyHash, *(PULONGLONG)CryptEncrypt, *(PULONGLONG)CryptDecrypt, *(PULONGLONG)CryptDestroyKey })
			g_crypt_code[i++] = p;
	}
	if (!(g_hResource = FindResource(NULL, SCRIPT_RESOURCE_NAME, RT_RCDATA)))
		g_hResource = FindResource(NULL, _T("E4847ED08866458F8DD35F94B37001C0"), RT_RCDATA);
#endif
	AHKModule();
	g_hInstance = hInstance;
	g_IconLarge = ExtractIconFromExecutable(NULL, -IDI_MAIN, 0, 0);
	g_IconSmall = ExtractIconFromExecutable(NULL, -IDI_MAIN, GetSystemMetrics(SM_CXSMICON), 0);
	g_FirstThreadID = GetCurrentThreadId();
	g_HistoryTickPrev = GetTickCount();
	g_TimeLastInputPhysical = GetTickCount();
	g_Debugger = new Debugger();
	g_script = new Script();
	g_clip = new Clipboard();
	g_MsgMonitor = new MsgMonitorList();
	InitializeCriticalSection(&g_Critical);
	InitializeCriticalSection(&g_CriticalRegExCache); // v1.0.45.04: Must be done early so that it's unconditional, so that DeleteCriticalSection() in the script destructor can also be unconditional (deleting when never initialized can crash, at least on Win 9x).
	InitializeCriticalSection(&g_CriticalTLSCallback);
	Object::sAnyPrototype = Object::CreateRootPrototypes();
	// v1.1.22+: This is done unconditionally, on startup, so that any attempts to read a drive
	// that has no media (and possibly other errors) won't cause the system to display an error
	// dialog that the script can't suppress.  This is known to affect floppy drives and some
	// but not all CD/DVD drives.  MSDN says: "Best practice is that all applications call the
	// process-wide SetErrorMode function with a parameter of SEM_FAILCRITICALERRORS at startup."
	// Note that in previous versions, this was done by the Drive/DriveGet commands and not
	// reverted afterward, so it affected all subsequent commands.
	SetErrorMode(SEM_FAILCRITICALERRORS);

	UpdateWorkingDir(); // Needed for the FileSelect() workaround.
	g_WorkingDirOrig = SimpleHeap::Alloc(g_WorkingDir.GetString()); // Needed by the Reload command.

	// Set defaults, to be overridden by command line args we receive:
	bool restart_mode = false;

	TCHAR *script_filespec = g_hResource ? SCRIPT_RESOURCE_SPEC : NULL; // Set default as "unspecified/omitted".

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
	for (i = 1; i < __argc; ++i) // Start at 1 because 0 contains the program name.
	{
		LPTSTR param = __targv[i]; // For performance and convenience.
		// Insist that switches be an exact match for the allowed values to cut down on ambiguity.
		// For example, if the user runs "CompiledScript.exe /find", we want /find to be considered
		// an input parameter for the script rather than a switch:
		if (!_tcsicmp(param, _T("/restart")))
			restart_mode = true;
		else if (!_tcsicmp(param, _T("/force")))
			g_ForceLaunch = true;
#ifndef AUTOHOTKEYSC // i.e. the following switch is recognized only by AutoHotkey.exe (especially since recognizing new switches in compiled scripts can break them, unlike AutoHotkey.exe).
		else if (!_tcsicmp(param, _T("/script"))) {
			if (g_hResource) {
				LPVOID data = LockResource(LoadResource(NULL, g_hResource));
				DWORD size = SizeofResource(NULL, g_hResource), old;
				VirtualProtect(data, size, PAGE_EXECUTE_READWRITE, &old);
				g_memset(data, 0, size);
			}
			script_filespec = NULL, g_hResource = NULL; // Override compiled script mode, otherwise no effect.
		}
		else if (script_filespec) // Compiled script mode.
			break;
		else if (!_tcsnicmp(param, _T("/ErrorStdOut"), 12) && (param[12] == '\0' || param[12] == '='))
			g_script->SetErrorStdOut(param[12] == '=' ? param + 13 : NULL);
		else if (!_tcsicmp(param, _T("/include")))
		{
			++i; // Consume the next parameter too, because it's associated with this one.
			if (i >= __argc // Missing the expected filename parameter.
				|| g_script->mCmdLineInclude) // Only one is supported, so abort if there's more.
				goto err;
			g_script->mCmdLineInclude = __targv[i];
		}
		else if (!_tcsicmp(param, _T("/validate")))
			g_script->mValidateThenExit = true;
		// DEPRECATED: /iLib
		else if (!_tcsicmp(param, _T("/iLib"))) // v1.0.47: Build an include-file so that ahk2exe can include library functions called by the script.
		{
			++i; // Consume the next parameter too, because it's associated with this one.
			if (i >= __argc) // Missing the expected filename parameter.
				goto err;
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
		else if (!g_Debugger->IsConnected() && !_tcsnicmp(param, _T("/Debug"), 6) && (param[6] == '\0' || param[6] == '='))
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
			if (!g_hResource)
			{
				script_filespec = param;  // The first unrecognized switch must be the script filespec, by design.
				++i; // Omit this from the "args" array.
			}
			break; // No more switches allowed after this point.
		}
	}

	if (Var *var = g_script->FindOrAddVar(_T("A_Args"), 6, VAR_DECLARE_GLOBAL))
	{
		// Store the remaining args in an array and assign it to "Args".
		// If there are no args, assign an empty array so that A_Args[1]
		// and A_Args.MaxIndex() don't cause an error.
		auto args = Array::FromArgV(__targv + i, __argc - i);
		if (!args)
			goto err;  // Realistically should never happen.
		var->AssignSkipAddRef(args);
	}
	else
		goto err;

	global_init(*g);  // Set defaults.

	// Set up the basics of the script:
	if (g_script->Init(*g, script_filespec, restart_mode, 0, _T("")) != OK)
		goto err;
	g_script->mInitFuncs = Array::Create();

	// Could use CreateMutex() but that seems pointless because we have to discover the
	// hWnd of the existing process so that we can close or restart it, so we would have
	// to do this check anyway, which serves both purposes.  Alt method is this:
	// Even if a 2nd instance is run with the /force switch and then a 3rd instance
	// is run without it, that 3rd instance should still be blocked because the
	// second created a 2nd handle to the mutex that won't be closed until the 2nd
	// instance terminates, so it should work ok:
	//CreateMutex(NULL, FALSE, script_filespec); // script_filespec seems a good choice for uniqueness.
	//if (!g_ForceLaunch && !restart_mode && GetLastError() == ERROR_ALREADY_EXISTS)

	UINT load_result = g_script->LoadFromFile(script_filespec);
	if (load_result == LOADING_FAILED) // Error during load (was already displayed by the function call).
		goto err;
	if (!load_result) // LoadFromFile() relies upon us to do this check.  No script was loaded or we're in /iLib mode, so nothing more to do.
		return 0;

	HWND w_existing = NULL;
	UserMessages reason_to_close_prior = (UserMessages)0;
	if (g_AllowOnlyOneInstance && !restart_mode && !g_ForceLaunch)
	{
		// Note: the title below must be constructed the same was as is done by our
		// CreateWindows(), which is why it's standardized in g_script->mMainWindowTitle:
		if (w_existing = FindWindow(g_WindowClassMain, g_script->mMainWindowTitle))
		{
			if (g_AllowOnlyOneInstance == SINGLE_INSTANCE_IGNORE)
				return 0;
			if (g_AllowOnlyOneInstance != SINGLE_INSTANCE_REPLACE)
				if (MsgBox(_T("An older instance of this script is already running.  Replace it with this")
					_T(" instance?\nNote: To avoid this message, see #SingleInstance in the help file.")
					, MB_YESNO, g_script->mFileName) == IDNO)
					return 0;
			// Otherwise:
			reason_to_close_prior = AHK_EXIT_BY_SINGLEINSTANCE;
		}
	}
	if (!reason_to_close_prior && restart_mode)
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
				{
					g_script->ExitApp(EXIT_SINGLEINSTANCE);
					return 0;
				}
				interval_count = 0;
			}
		}
		// Give it a small amount of additional time to completely terminate, even though
		// its main window has already been destroyed:
		Sleep(100);
	}

	// Create all our windows and the tray icon.  This is done after all other chances
	// to return early due to an error have passed, above.
	if (g_script->CreateWindows() != OK)
		goto err;

	// store main thread window as first item
	g_ahkThreads[0] = { g_hWnd,g_script,((PMYTEB)NtCurrentTeb())->ThreadLocalStoragePointer,g_FirstThreadID };

	ChangeWindowMessageFilter(WM_COMMNOTIFY, MSGFLT_ADD);

	// At this point, it is nearly certain that the script will be executed.

	// v1.0.48.04: Turn off buffering on stdout so that "FileAppend, Text, *" will write text immediately
	// rather than lazily. This helps debugging, IPC, and other uses, probably with relatively little
	// impact on performance given the OS's built-in caching.  I looked at the source code for setvbuf()
	// and it seems like it should execute very quickly.  Code size seems to be about 75 bytes.
	setvbuf(stdout, NULL, _IONBF, 0); // Must be done PRIOR to writing anything to stdout.

	if (g_MaxHistoryKeys && (g_KeyHistory = (KeyHistoryItem *)malloc(g_MaxHistoryKeys * sizeof(KeyHistoryItem))))
		ZeroMemory(g_KeyHistory, g_MaxHistoryKeys * sizeof(KeyHistoryItem)); // Must be zeroed.
	//else leave it NULL as it was initialized in globaldata.

#ifdef CONFIG_DEBUGGER
	// Initiate debug session now if applicable.
	if (!g_script->mEncrypt && !g_DebuggerHost.IsEmpty() && g_Debugger->Connect(g_DebuggerHost, g_DebuggerPort) == DEBUGGER_E_OK)
	{
		g_Debugger->Break();
	}
#endif

	// Activate the hotkeys, hotstrings, and any hooks that are required prior to executing the
	// top part (the auto-execute part) of the script so that they will be in effect even if the
	// top part is something that's very involved and requires user interaction:
	Hotkey::ManifestAllHotkeysHotstringsHooks(); // We want these active now in case auto-execute never returns (e.g. loop)
	g_script->mIsReadyToExecute = true; // This is done only after the above to support error reporting in Hotkey.cpp.
	g_HSSameLineAction = false; // `#Hotstring X` should not affect Hotstring().
	g_SuspendExempt = false; // #SuspendExempt should not affect Hotkey()/Hotstring().

	return MainExecuteScript();
err:
	g_script->TerminateApp(EXIT_CRITICAL, CRITICAL_ERROR);
	return CRITICAL_ERROR;
}


int MainExecuteScript()
{

#ifndef _DEBUG
	__try
#endif
	{
		// Run the auto-execute part at the top of the script (this call might never return):
		if (!g_script->AutoExecSection()) // Can't run script at all. Due to rarity, just abort.
			return CRITICAL_ERROR;
		// REMEMBER: The call above will never return if one of the following happens:
		// 1) The AutoExec section never finishes (e.g. infinite loop).
		// 2) The AutoExec function uses the Exit or ExitApp command to terminate the script.
		// 3) The script isn't persistent and its last line is reached (in which case an ExitApp is implicit).

		// Call it in this special mode to kick off the main event loop.
		// Be sure to pass something >0 for the first param or it will
		// return (and we never want this to return):
		MsgSleep(SLEEP_INTERVAL, WAIT_FOR_MESSAGES);
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
	return 0; // Never executed; avoids compiler warning.
}
#endif // !_USRDLL

unsigned __stdcall ThreadMain(LPTSTR *data)
{
	size_t len = (size_t)data[0];
	TextMem::Buffer buf(malloc((len + MAX_INTEGER_LENGTH + 1) * sizeof(TCHAR)));
	if (!buf.mBuffer)
		return CRITICAL_ERROR;
	LPTSTR lpScript = (LPTSTR)buf.mBuffer + MAX_INTEGER_LENGTH, lpTitle = _T("AutoHotkey");
	int argc = 0;
	LPTSTR* argv = NULL;
	TCHAR filepath[MAX_PATH];
	DWORD encrypt = g_ahkThreads[0].Script ? g_ahkThreads[0].Script->mEncrypt : 0;

	if (data[3])
		lpScript += _stprintf(lpTitle = lpScript, _T("%s"), data[3]) + 1;
	if (data[2])
		argv = CommandLineToArgvW(data[2], &argc);
	if (data[1])
		_tcscpy(lpScript, data[1]), len -= lpScript - (LPTSTR)buf.mBuffer - MAX_INTEGER_LENGTH;
	else
		lpScript = _T("Persistent"), len = 11;

	auto lps = lpScript + crypt::linear_congruent_generator((int)(ULONG_PTR)lpScript & 8);
	_stprintf(filepath, _T("*THREAD%u?%p#%zu.AHK"), g_MainThreadID, encrypt ? lps : lpScript, encrypt ? 0 : len * sizeof(TCHAR));
	sntprintf((LPTSTR)buf.mBuffer, MAX_INTEGER_LENGTH, _T("ahk%d"), GetCurrentThreadId());
	HANDLE hEvent = OpenEvent(EVENT_MODIFY_STATE, true, buf);

	InitializeCriticalSection(&g_CriticalRegExCache); // v1.0.45.04: Must be done early so that it's unconditional, so that DeleteCriticalSection() in the script destructor can also be unconditional (deleting when never initialized can crash, at least on Win 9x).
	InitializeCriticalSection(&g_CriticalTLSCallback);

	for (;;)
	{
		g_Debugger = new Debugger();
		g_script = new Script();
		g_clip = new Clipboard();
		g_MsgMonitor = new MsgMonitorList();
		Object::sAnyPrototype = Object::CreateRootPrototypes();
		g_script->mEncrypt = encrypt;

		// v1.1.22+: This is done unconditionally, on startup, so that any attempts to read a drive
		// that has no media (and possibly other errors) won't cause the system to display an error
		// dialog that the script can't suppress.  This is known to affect floppy drives and some
		// but not all CD/DVD drives.  MSDN says: "Best practice is that all applications call the
		// process-wide SetErrorMode function with a parameter of SEM_FAILCRITICALERRORS at startup."
		// Note that in previous versions, this was done by the Drive/DriveGet commands and not
		// reverted afterward, so it affected all subsequent commands.
		SetErrorMode(SEM_FAILCRITICALERRORS);

		UpdateWorkingDir(); // Needed for the FileSelect() workaround.
		g_WorkingDirOrig = SimpleHeap::Malloc(const_cast<LPTSTR>(g_WorkingDir.GetString())); // Needed by the Reload command.

		// Set defaults, to be overridden by command line args we receive:
		bool restart_mode = false;

		for (int i = 0; i < argc; ++i) // Start at 0 it does not contain the program name.
		{
			LPTSTR param = argv[i]; // For performance and convenience.
			// Insist that switches be an exact match for the allowed values to cut down on ambiguity.
			// For example, if the user runs "CompiledScript.exe /find", we want /find to be considered
			// an input parameter for the script rather than a switch:
			if (!_tcsicmp(param, _T("/ErrorStdOut")))
				g_script->SetErrorStdOut(param[12] == '=' ? param + 13 : NULL);
			else if (!_tcsicmp(param, _T("/validate")))
				g_script->mValidateThenExit = true;
			else if (!_tcsicmp(param, _T("/iLib"))) // v1.0.47: Build an include-file so that ahk2exe can include library functions called by the script.
			{
				++i; // Consume the next parameter too, because it's associated with this one.
				if (i >= argc) // Missing the expected filename parameter.
					goto err;
				// For performance and simplicity, open/create the file unconditionally and keep it open until exit.
				g_script->mValidateThenExit = true;
			}
			else if (!_tcsnicmp(param, _T("/CP"), 3)) // /CPnnn
			{
				// Default codepage for the script file, NOT the default for commands used by it.
				g_DefaultScriptCodepage = ATOU(param + 3);
			}
#ifdef CONFIG_DEBUGGER
			// Allow a debug session to be initiated by command-line.
			else if (!encrypt && !g_Debugger->IsConnected() && !_tcsnicmp(param, _T("/Debug"), 6) && (param[6] == '\0' || param[6] == '='))
			{
				if (param[6] == '=')
				{
					param += 7;

					LPTSTR c = _tcsrchr(param, ':');

					if (c)
					{
						StringTCharToChar(param, g_DebuggerHost, (int)(c - param));
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
			else if (!g_script->mEncrypt && !_tcsnicmp(param, _T("/NoDebug"), 8)) {
				g_script->mEncrypt = 1;
				_stprintf(filepath, _T("*THREAD%u?%p#%zu.AHK"), g_MainThreadID, lps, (size_t)0);
			}
			else // since this is not a recognized switch, the end of the [Switches] section has been reached (by design).
			{
				break; // No more switches allowed after this point.
			}
		}

		if (Var *var = g_script->FindOrAddVar(_T("A_Args"), 6, VAR_DECLARE_GLOBAL))
		{
			// Store the remaining args in an array and assign it to "A_Args".
			// If there are no args, assign an empty array so that A_Args[1]
			// and A_Args.Length() don't cause an error.
			auto args = Array::FromArgV(argv, argc);
			if (!args)
				goto err;
			var->AssignSkipAddRef(args);
		}
		else
			goto err;

		global_init(*g);  // Set defaults prior to the below, since below might override them for AutoIt2 scripts.
		g_NoTrayIcon = true;
		// Set up the basics of the script:
		if (g_script->Init(*g, filepath, 0, g_hInstance, lpTitle) != OK) // Set up the basics of the script, using the above.
			goto err;
		g_script->mInitFuncs = Array::Create();

		//if (nameHinstanceP.istext)
		//	GetCurrentDirectory(MAX_PATH, g_script->mFileDir);
		// Could use CreateMutex() but that seems pointless because we have to discover the
		// hWnd of the existing process so that we can close or restart it, so we would have
		// to do this check anyway, which serves both purposes.  Alt method is this:
		// Even if a 2nd instance is run with the /force switch and then a 3rd instance
		// is run without it, that 3rd instance should still be blocked because the
		// second created a 2nd handle to the mutex that won't be closed until the 2nd
		// instance terminates, so it should work ok:
		//CreateMutex(NULL, FALSE, script_filespec); // script_filespec seems a good choice for uniqueness.
		//if (!g_ForceLaunch && !restart_mode && GetLastError() == ERROR_ALREADY_EXISTS)

		LineNumberType load_result = g_script->LoadFromText(lpScript, filepath);
		if (load_result == LOADING_FAILED) // Error during load (was already displayed by the function call).
			goto err;  // Should return this value because PostQuitMessage() also uses it.
		if (!load_result) // LoadFromFile() relies upon us to do this check.  No lines were loaded, so we're done.
		{
			g_script->TerminateApp(EXIT_EXIT, g_ExitCode = 0);
			goto exit;
		}
		if (*lpTitle)
			g_script->mFileName = lpTitle;
		// Create all our windows and the tray icon.  This is done after all other chances
		// to return early due to an error have passed, above.
		if (g_script->CreateWindows() != OK)
			goto err;

		// save g_hWnd in threads global array used for exiting threads, debugger and probably more
		EnterCriticalSection(&g_Critical);
		for (int i = 1;; i++)
		{
			if (i == MAX_AHK_THREADS)
				goto err;
			if (!IsWindow(g_ahkThreads[i].Hwnd))
			{
				g_ahkThreads[i] = { g_hWnd,g_script,((PMYTEB)NtCurrentTeb())->ThreadLocalStoragePointer,g_MainThreadID };
				break;
			}
		}
		LeaveCriticalSection(&g_Critical);
		// Set AutoHotkey.dll its main window (ahk_class AutoHotkey) bottom so it does not receive Reload or ExitApp Message send e.g. when Reload message is sent.
		//SetWindowPos(g_hWnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOCOPYBITS | SWP_NOMOVE | SWP_NOREDRAW | SWP_NOSENDCHANGING | SWP_NOSIZE);

		// At this point, it is nearly certain that the script will be executed.

		// v1.0.48.04: Turn off buffering on stdout so that "FileAppend, Text, *" will write text immediately
		// rather than lazily. This helps debugging, IPC, and other uses, probably with relatively little
		// impact on performance given the OS's built-in caching.  I looked at the source code for setvbuf()
		// and it seems like it should execute very quickly.  Code size seems to be about 75 bytes.
		setvbuf(stdout, NULL, _IONBF, 0); // Must be done PRIOR to writing anything to stdout.

		//if (g_MaxHistoryKeys && !g_KeyHistory && (g_KeyHistory = (KeyHistoryItem *)malloc(g_MaxHistoryKeys * sizeof(KeyHistoryItem))))
		//	ZeroMemory(g_KeyHistory, g_MaxHistoryKeys * sizeof(KeyHistoryItem)); // Must be zeroed.
		//else leave it NULL as it was initialized in globaldata.

#ifdef CONFIG_DEBUGGER
	// Initiate debug session now if applicable.
		if (!g_script->mEncrypt && !g_DebuggerHost.IsEmpty() && g_Debugger->Connect(g_DebuggerHost, g_DebuggerPort) == DEBUGGER_E_OK)
		{
			g_Debugger->Break();
		}
#endif


		// Activate the hotkeys, hotstrings, and any hooks that are required prior to executing the
		// top part (the auto-execute part) of the script so that they will be in effect even if the
		// top part is something that's very involved and requires user interaction:
		Hotkey::ManifestAllHotkeysHotstringsHooks(); // We want these active now in case auto-execute never returns (e.g. loop)
		g_HSSameLineAction = false; // `#Hotstring X` should not affect Hotstring().
		g_SuspendExempt = false; // #SuspendExempt should not affect Hotkey()/Hotstring().
		g_script->mIsReadyToExecute = true; // This is done only after the above to support error reporting in Hotkey.cpp.

		if (hEvent) {
			data[4] = (LPTSTR)(UINT_PTR)g_MainThreadID;
			SetEvent(hEvent);
			CloseHandle(hEvent);
			hEvent = NULL;
		}
		if (ThreadExecuteScript() != CONDITION_TRUE)
			break;
	}
	goto exit;
err:
	if (hEvent) {
		SetEvent(hEvent);
		CloseHandle(hEvent);
		hEvent = NULL;
	}
	g_script->TerminateApp(EXIT_EXIT, g_ExitCode = CRITICAL_ERROR);
exit:
	if (argv)
		LocalFree(argv); // free memory allocated by CommandLineToArgvW
	return g_ExitCode;
}


ResultType ThreadExecuteScript()
{
	
#ifndef _DEBUG
	__try
#endif
	{

		// Run the auto-execute part at the top of the script (this call might never return):
		ResultType result = g_script->AutoExecSection();
		if (!result) // Can't run script at all. Due to rarity, just abort.
		{
			g_ExitCode = CRITICAL_ERROR;
			return CRITICAL_ERROR;
		}
		else if (result != OK || g_Reloading) {
			g_script->TerminateApp(EXIT_EXIT, g_ExitCode = 0);
			if (g_Reloading) {
				g_Reloading = false;
				return CONDITION_TRUE;
			}
			return OK;
		}
		// REMEMBER: The call above will never return if one of the following happens:
		// 1) The AutoExec section never finishes (e.g. infinite loop).
		// 2) The AutoExec function uses the Exit or ExitApp command to terminate the script.
		// 3) The script isn't persistent and its last line is reached (in which case an ExitApp is implicit).

		// Call it in this special mode to kick off the main event loop.
		// Be sure to pass something >0 for the first param or it will
		// return (and we never want this to return):
		MsgSleep(SLEEP_INTERVAL, WAIT_FOR_MESSAGES);
		g_script->TerminateApp(EXIT_EXIT, g_ExitCode = 0);
		if (g_Reloading) {
			g_Reloading = false;
			return CONDITION_TRUE;
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
		g_ExitCode = CRITICAL_ERROR;
	}
#endif
	return OK;
}