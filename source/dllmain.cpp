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
#ifdef _USRDLL
#include "globaldata.h"
#include "COMServer.h"
//#include "xdlldata.h"

CAutoHotkeyModule _AtlModule;

extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	switch (dwReason)
	{
	case DLL_PROCESS_ATTACH: {
		g_hInstance = hInstance;
		g_hMemoryModule = (HMODULE)lpReserved;
		InitializeCriticalSection(&g_Critical);
		g_IconLarge = ExtractIconFromExecutable(NULL, -IDI_MAIN, 0, 0);
		g_IconSmall = ExtractIconFromExecutable(NULL, -IDI_MAIN, GetSystemMetrics(SM_CXSMICON), 0);
		int i = 0;
		for (auto p : { *(PULONGLONG)CryptHashData, *(PULONGLONG)CryptDeriveKey, *(PULONGLONG)CryptDestroyHash, *(PULONGLONG)CryptEncrypt, *(PULONGLONG)CryptDecrypt, *(PULONGLONG)CryptDestroyKey })
			g_crypt_code[i++] = p;
		break;
	}
	case DLL_THREAD_ATTACH:
		break;

	case DLL_PROCESS_DETACH:
	case DLL_THREAD_DETACH: {
		if (g_script)
			g_script->TerminateApp(ExitReasons::EXIT_EXIT, 0);
		if (dwReason == DLL_PROCESS_DETACH) {
			if (g_hWinAPI)
			{
				free(g_hWinAPI), g_hWinAPI = NULL;;
				free(g_hWinAPIlowercase), g_hWinAPIlowercase = NULL;
			}
			DeleteCriticalSection(&g_Critical);
			if (g_IconLarge)
				DestroyIcon(g_IconLarge), g_IconLarge = NULL;
			if (g_IconSmall)
				DestroyIcon(g_IconSmall), g_IconSmall = NULL;
		}
		break;
	}
	}
	return _AtlModule.DllMain(dwReason, lpReserved);
}

_Use_decl_annotations_
STDAPI DllCanUnloadNow(void)
{
	return _AtlModule.DllCanUnloadNow();
}

_Use_decl_annotations_
STDAPI DllGetClassObject(_In_ REFCLSID rclsid, _In_ REFIID riid, _Outptr_ LPVOID* ppv)
{
	return _AtlModule.DllGetClassObject(rclsid, riid, ppv);
}

_Use_decl_annotations_
STDAPI DllRegisterServer(void)
{
	return _AtlModule.DllRegisterServer();
}

_Use_decl_annotations_
STDAPI DllUnregisterServer(void)
{
	return _AtlModule.DllUnregisterServer();
}

STDAPI DllInstall(BOOL bInstall, _In_opt_  LPCWSTR pszCmdLine)
{
	HRESULT hr = E_FAIL;
	static const wchar_t szUserSwitch[] = L"user";

	if (pszCmdLine != nullptr)
	{
		if (_wcsnicmp(pszCmdLine, szUserSwitch, _countof(szUserSwitch)) == 0)
		{
			ATL::AtlSetPerUserRegistration(true);
		}
	}

	if (bInstall)
	{
		hr = DllRegisterServer();
		if (FAILED(hr))
		{
			DllUnregisterServer();
		}
	}
	else
	{
		hr = DllUnregisterServer();
	}

	return hr;
}

#endif