#include "stdafx.h" // pre-compiled headers
#include "defines.h"
#include "globaldata.h"
#include "script.h"
#include "application.h"

#include "TextIO.h"
#include "script_object.h"
#include "script_func_impl.h"
#include "script_com.h"

#include "LiteZip.h"
#include "MemoryModule.h"
#include "exports.h"


////////////////////////////
// AUTOHOTKEY_H FUNCTIONS //
////////////////////////////
MdType ConvertDllArgType(LPTSTR aBuf, bool &aPassedByAddr);
BIF_DECL(BIF_Cast)
{
	MdType types[2];
	bool passed_by_addr[2];
	int i = 0;
	for (; i < 2; i++)
	{
		types[i] = ConvertDllArgType(TokenToString(*aParam[i * 2]), passed_by_addr[i]);
		if (types[i] == MdType::Void)
			_o_throw_value(ERR_PARAM_INVALID);
	}
	ExprTokenType aValue;
	void *ptr;
	if (TokenToDoubleOrInt64(*aParam[1], aValue) != OK)
		_o_throw_param(1, _T("Number"));
	if (passed_by_addr[0])
	{
		ptr = (void *)TokenToInt64(aValue);
		if ((size_t)ptr < 65536)
			_o_throw_value(ERR_PARAM2_INVALID);
	}
	else
		ptr = &aResultToken.value_int64;
	for (i = 0; i < 2; i++)
	{
		auto tp = types[i];
		if (passed_by_addr[i])
		{
			TypedPtrToToken(tp, ptr, aResultToken);
			ptr = &aResultToken.value_int64;
			aValue.CopyExprFrom(aResultToken);
			continue;
		}
		aResultToken.value_int64 = 0;
		SetValueOfTypeAtPtr(tp, ptr, aValue, aResultToken);
		if (tp == MdType::Float32)
			aValue.SetValue(*(float *)ptr);
		else
		{
			if (tp == MdType::Float64)
				aValue.symbol = SYM_FLOAT;
			else aValue.symbol = SYM_INTEGER;
			aValue.value_int64 = aResultToken.value_int64;
		}
		aResultToken.symbol = aValue.symbol;
	}
}

BIF_DECL(BIF_CryptAES)
{
	TCHAR *pwd;
	size_t ptr = 0, size = 0;
	BOOL encrypt;
	DWORD sid;
	LPVOID data;
	if (auto obj = TokenToObject(*aParam[0]))
	{
		GetBufferObjectPtr(aResultToken, obj, ptr, size);
		if (aResultToken.Exited())
			return;
		if (ParamIndexIsOmitted(1))
			_o_throw_param(1);
	}
	else {
		if (ParamIndexIsOmitted(2))
			_o_throw_param(2);
		if (TokenIsPureNumeric(*aParam[0]))
			ptr = (size_t)ParamIndexToInt64(0);
		else
			ptr = (size_t)ParamIndexToString(0);
		size = (size_t)ParamIndexToInt64(1);
		aParam++, aParamCount--;
	}
	if (ptr < 65536 || !size)
		_o_throw_value(ERR_INVALID_VALUE);
	pwd = TokenToString(*aParam[1]);
	encrypt = ParamIndexToOptionalBOOL(2, TRUE);
	sid = ParamIndexToOptionalInt(3, 256);
	if (!(data = malloc(size + 16)))
		_o_throw_oom;
	memcpy(data, (const void *)ptr, size);
	if (size_t newsize = CryptAES(data, (DWORD)size, pwd, encrypt, sid))
		_o_return(BufferObject::Create(data, newsize));
	free(data);
	_o_throw(ERR_FAILED);
}

BIF_DECL(BIF_ResourceLoadLibrary)
{
	auto resname = ParamIndexToString(0);
	auto inst = g_hInstance;
	HRSRC hRes;
	void *data, *buf = nullptr;
	DWORD size;
	if (!*resname)
		_o_throw_value(ERR_PARAM1_INVALID);

	if (!(hRes = FindResource(g_hInstance, ParamIndexToString(0), RT_RCDATA)))
		hRes = FindResource(inst = NULL, ParamIndexToString(0), RT_RCDATA);

	if (!(hRes
		&& (size = SizeofResource(inst, hRes))
		&& (data = LockResource(LoadResource(inst, hRes)))))
		_o_throw_win32();
	if (*(unsigned int *)data == 0x04034b50)
		if (auto aSizeDeCompressed = DecompressBuffer(data, buf, size))
			data = buf, size = aSizeDeCompressed;

	aResultToken.SetValue((UINT_PTR)MemoryLoadLibrary(data, size));
	free(buf);
}

BIF_DECL(BIF_MemoryCallEntryPoint)
{
	auto mod = (HMEMORYMODULE)TokenToInt64(*aParam[0]);
	if (!mod)
		_o_throw_value(ERR_PARAM1_INVALID);
	_o_return((UINT_PTR)MemoryCallEntryPoint(mod, ParamIndexToOptionalStringDef(1, _T(""))));
}

static void FileRead(ResultToken &aResultToken, LPTSTR aPath, void *&aData, size_t &aSize)
{
	HANDLE hfile = CreateFile(aPath, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING
		, FILE_FLAG_SEQUENTIAL_SCAN, NULL);
	if (hfile == INVALID_HANDLE_VALUE)
		_o_throw_win32(g->LastError = GetLastError());
	aSize = (size_t)GetFileSize64(hfile);
	if (aSize == ULLONG_MAX) // GetFileSize64() failed.
	{
		g->LastError = GetLastError();
		CloseHandle(hfile);
		_o_throw_win32(g->LastError);
	}
	if (!(aData = malloc(aSize)))
	{
		CloseHandle(hfile);
		return (void)aResultToken.MemoryError();
	}
	DWORD err = 0;
	if (!ReadFile(hfile, aData, (DWORD)aSize, (LPDWORD)&aSize, NULL))
		err = GetLastError();
	CloseHandle(hfile);
	if (!err)
		return;
	free(aData);
	_o_throw_win32(err);
}

BIF_DECL(BIF_MemoryLoadLibrary)
{
	void *data, *buf = nullptr;
	size_t size = 0;
	if (auto obj = TokenToObject(*aParam[0]))
	{
		GetBufferObjectPtr(aResultToken, obj, (size_t&)data, size);
		if (aResultToken.Exited())
			return;
	}
	else if (TokenIsPureNumeric(*aParam[0]))
	{
		data = (void *)TokenToInt64(*aParam[0]);
		size = (size_t)TokenToInt64(*aParam[1]);
	}
	else
	{
		auto path = TokenToString(*aParam[0]);
		if (!*path)
			_o_throw_value(ERR_PARAM1_INVALID);
		FileRead(aResultToken, path, data, size);
		if (aResultToken.Exited())
			return;
		buf = data;
	}
	if (!data)
		_o_throw_value(ERR_PARAM1_INVALID);
	aResultToken.SetValue((UINT_PTR)MemoryLoadLibraryEx(data, size,
		MemoryDefaultAlloc, MemoryDefaultFree,
		(CustomLoadLibraryFunc)ParamIndexToOptionalIntPtr(2, (UINT_PTR)MemoryDefaultLoadLibrary),
		(CustomGetProcAddressFunc)ParamIndexToOptionalIntPtr(3, (UINT_PTR)MemoryDefaultGetProcAddress),
		(CustomFreeLibraryFunc)ParamIndexToOptionalIntPtr(4, (UINT_PTR)MemoryDefaultFreeLibrary),
		NULL
	));
	free(buf);
}

BIF_DECL(BIF_MemoryGetProcAddress)
{
	auto mod = (HMEMORYMODULE)TokenToInt64(*aParam[0]);
	auto procname = TokenToString(*aParam[1]);
	char *name = nullptr, *tofree = nullptr;
	if (!mod)
		_o_throw_value(ERR_PARAM1_INVALID);
	if (*procname)
	{
#ifdef _UNICODE
		name = tofree = CStringCharFromWChar(procname).DetachBuffer();
#else
		name = procname;
#endif
	}
	else
	{
		if (TokenIsPureNumeric(*aParam[1]) == SYM_INTEGER)
			if (HIWORD(name = (char *)TokenToInt64(*aParam[1])))
				name = nullptr;
		if (!name)
			_o_throw_value(ERR_PARAM2_INVALID);
	}
	aResultToken.SetValue((__int64)MemoryGetProcAddress(mod, name));
}

BIF_DECL(BIF_MemoryFreeLibrary)
{
	auto mod = (HMEMORYMODULE)TokenToInt64(*aParam[0]);
	if (!mod)
		_o_throw_value(ERR_PARAM1_INVALID);
	MemoryFreeLibrary(mod);
}

BIF_DECL(BIF_MemoryFindResource)
{
	auto mod = (HMEMORYMODULE)TokenToInt64(*aParam[0]);
	if (!mod)
		_o_throw_value(ERR_PARAM1_INVALID);
	LPCTSTR param[2] = { 0 };
	for (int i = 1; i < 3; i++)
	{
		auto &token = *aParam[i];
		LPCTSTR t = nullptr;
		if (TokenIsPureNumeric(token) == SYM_INTEGER)
		{
			if (!IS_INTRESOURCE(t = (LPCTSTR)TokenToInt64(token)))
				t = nullptr;
		}
		else if (!*(t = TokenToString(token)))
			t = nullptr;
		if (!(param[i - 1] = t))
			_o_throw_value(ERR_PARAM_INVALID);
	}
	HMEMORYRSRC hres = MemoryFindResourceEx(mod, param[0], param[1],
		(WORD)ParamIndexToOptionalInt(3, MAKELANGID(LANG_NEUTRAL, SUBLANG_NEUTRAL)));
	aResultToken.SetValue((__int64)hres);
}

BIF_DECL(BIF_MemorySizeOfResource)
{
	auto mod = (HMEMORYMODULE)TokenToInt64(*aParam[0]);
	if (!mod)
		_o_throw_value(ERR_PARAM1_INVALID);
	auto hres = (HMEMORYRSRC)TokenToInt64(*aParam[1]);
	if (!hres)
		_o_throw_value(ERR_PARAM2_INVALID);
	_o_return(MemorySizeOfResource(mod, hres));
}

BIF_DECL(BIF_MemoryLoadResource)
{
	auto mod = (HMEMORYMODULE)TokenToInt64(*aParam[0]);
	if (!mod)
		_o_throw_value(ERR_PARAM1_INVALID);
	auto hres = (HMEMORYRSRC)TokenToInt64(*aParam[1]);
	if (!hres)
		_o_throw_value(ERR_PARAM2_INVALID);
	_o_return((UINT_PTR)MemoryLoadResource(mod, hres));
}

BIF_DECL(BIF_MemoryLoadString)
{
	auto mod = (HMEMORYMODULE)TokenToInt64(*aParam[0]);
	if (!mod)
		_o_throw_value(ERR_PARAM1_INVALID);
	auto str = MemoryLoadStringEx(mod, (UINT)TokenToInt64(*aParam[1]), (WORD)ParamIndexToOptionalInt(2, 0));
	if (str)
		aResultToken.AcceptMem(str, _tcslen(str));
}

BIF_DECL(BIF_Swap)
{
	if (aParam[0]->symbol != SYM_VAR || aParam[1]->symbol != SYM_VAR)
		_o_throw_value(ERR_INVALID_ARG_TYPE);
	Var &a = *aParam[0]->var->ResolveAlias(), &b = *aParam[1]->var->ResolveAlias();
	swap(a, b);
	swap(a.mName, b.mName);
	swap(a.mType, b.mType);
	swap(a.mScope, b.mScope);
}

static void ZipError(ResultToken &aResultToken, DWORD aErrCode)
{
	TCHAR buf[100];
	ZipFormatMessage(aErrCode, buf, _countof(buf));
	_o_throw(buf);
}

#define _o_throw_zip(code)	return ZipError(aResultToken, code)
BIF_DECL(BIF_UnZip)
{
	IObject *obj;
	size_t	ptr = 0, size;
	DWORD	err;
	HZIP	huz;
	LPTSTR	path = nullptr;
	if (obj = TokenToObject(*aParam[0]))
	{
		GetBufferObjectPtr(aResultToken, obj, ptr, size);
		if (aResultToken.Exited())
			return;
		if (ptr < 65536 || !size)
			_o_throw_value(ERR_PARAM1_INVALID);
	}
	else if (TokenIsPureNumeric(*aParam[0]))
	{
		if (!TokenIsNumeric(*aParam[1]))
			_o_throw_param(1, _T("Integer"));
		if (ptr = (size_t)TokenToInt64(*aParam[0]))
		{
			size = (size_t)TokenToInt64(*aParam[1]);
			if (ptr < 65536)
				_o_throw_value(ERR_PARAM1_INVALID);
			if (!size)
				_o_throw_value(ERR_PARAM2_INVALID);
		}
		aParam++;
		aParamCount--;
	}
	else
		path = TokenToString(*aParam[0]);
	if (ParamIndexIsOmitted(1))
		_o_throw_param(!path && !obj ? 2 : 1);

	auto codepage = (UINT)ParamIndexToOptionalInt(5, 0);
	CStringCharFromTChar pw(ParamIndexToOptionalString(4));
	if (path)
		err = UnzipOpenFileW(&huz, path, pw, codepage);
	else
		err = UnzipOpenBuffer(&huz, (void *)ptr, (ULONGLONG)size, pw, codepage);
	if (err)
		goto error;

	ZIPENTRY ze;
	LPTSTR aDir = TokenToString(*aParam[1], NULL, &size);
	TCHAR aTargetDir[MAX_PATH] = { 0 };
	if (size > MAX_PATH - 2)
	{
		UnzipClose(huz);
		_o_throw(_T("Path too long."));
	}
	_tcscpy(aTargetDir, aDir);
	if (size && aTargetDir[size - 1] != '\\')
		aTargetDir[size++] = '\\';

	if (ParamIndexIsOmitted(2))
	{
	extractall:
		ULONGLONG	numitems;
		ze.Index = (ULONGLONG)-1;
		if (err = UnzipGetItem(huz, &ze))
			goto errorclose;
		numitems = ze.Index;
		for (ze.Index = 0; ze.Index < numitems; ze.Index++)
		{
			if ((err = UnzipGetItem(huz, &ze)))
				goto errorclose;
			_tcscpy(aTargetDir + size, ze.Name + 1);
			if (err = UnzipItemToFile(huz, aTargetDir, &ze))
				goto errorclose;
		}
	}
	else
	{
		if (TokenIsPureNumeric(*aParam[2]))
		{
			ze.Index = (ULONGLONG)TokenToInt64(*aParam[2]);
			err = UnzipGetItem(huz, &ze);
		}
		else
		{
			auto item = TokenToString(*aParam[2]);
			if (!*item)
				goto extractall;
			if (_tcslen(item) > MAX_PATH - 2)
			{
				UnzipClose(huz);
				_o_throw(_T("Path too long."));
			}
			_tcscpy(ze.Name, item);
			err = UnzipFindItem(huz, &ze, 1);
		}
		if (err) goto errorclose;
		_tcscpy(aTargetDir + size, !ParamIndexIsOmittedOrEmpty(3) ? TokenToString(*aParam[3]) : ze.Name + 1);
		if (err = UnzipItemToFile(huz, aTargetDir, &ze))
			goto errorclose;
	}
	UnzipClose(huz);
	return;
errorclose:
	UnzipClose(huz);
error:
	_o_throw_zip(err);
}

BIF_DECL(BIF_UnZipBuffer)
{
	IObject *obj;
	size_t	ptr = 0, size;
	DWORD	err;
	HZIP	huz;
	LPTSTR	path = nullptr;
	if (obj = TokenToObject(*aParam[0]))
	{
		GetBufferObjectPtr(aResultToken, obj, ptr, size);
		if (aResultToken.Exited())
			return;
		if (ptr < 65536 || !size)
			_o_throw_value(ERR_PARAM1_INVALID);
	}
	else if (TokenIsPureNumeric(*aParam[0]))
	{
		if (!TokenIsNumeric(*aParam[1]))
			_o_throw_param(1, _T("Integer"));
		if (ptr = (size_t)TokenToInt64(*aParam[0]))
		{
			size = (size_t)TokenToInt64(*aParam[1]);
			if (ptr < 65536)
				_o_throw_value(ERR_PARAM1_INVALID);
			if (!size)
				_o_throw_value(ERR_PARAM2_INVALID);
		}
		aParam++;
		aParamCount--;
	}
	else
		path = TokenToString(*aParam[0]);
	if (ParamIndexIsOmittedOrEmpty(1))
		_o_throw_param(!path && !obj ? 2 : 1);

	auto codepage = (UINT)ParamIndexToOptionalInt(3, 0);
	CStringCharFromTChar pw(ParamIndexToOptionalString(2));
	if (path)
		err = UnzipOpenFileW(&huz, path, pw, codepage);
	else
		err = UnzipOpenBuffer(&huz, (void *)ptr, (ULONGLONG)size, pw, codepage);
	if (err)
		goto error;

	ZIPENTRY	ze;
	if (TokenIsPureNumeric(*aParam[1]))
	{
		ze.Index = (ULONGLONG)TokenToInt64(*aParam[1]);
		err = UnzipGetItem(huz, &ze);
	}
	else
	{
		auto item = TokenToString(*aParam[1]);
		if (_tcslen(item) > MAX_PATH - 2)
		{
			UnzipClose(huz);
			_o_throw(_T("Path too long."));
		}
		_tcscpy(ze.Name, item);
		err = UnzipFindItem(huz, &ze, 1);
	}
	if (err) goto errorclose;
	auto buf = malloc((size_t)ze.UncompressedSize);
	if (!buf)
	{
		UnzipClose(huz);
		_o_throw_oom;
	}
	if (err = UnzipItemToBuffer(huz, buf, &ze))
	{
		free(buf);
		goto errorclose;
	}
	UnzipClose(huz);
	_o_return(BufferObject::Create(buf, (size_t)ze.UncompressedSize));
errorclose:
	UnzipClose(huz);
error:
	_o_throw_zip(err);
}

BIF_DECL(BIF_UnZipRawMemory)
{
	size_t ptr, size;
	LPTSTR pwd;
	LPVOID buf;

	if (auto obj = TokenToObject(*aParam[0]))
	{
		GetBufferObjectPtr(aResultToken, obj, ptr, size);
		if (aResultToken.Exited())
			return;
		pwd = ParamIndexIsOmittedOrEmpty(1) ? NULL : TokenToString(*aParam[1]);
	}
	else
	{
		if (ParamIndexIsOmitted(1))
			_o_throw_param(1);
		ptr = (size_t)TokenToInt64(*aParam[0]);
		size = (size_t)TokenToInt64(*aParam[1]);
		pwd = ParamIndexIsOmittedOrEmpty(2) ? NULL : TokenToString(*aParam[2]);
	}
	if (ptr < 65536 || !size)
		_o_throw_value(ERR_PARAM_INVALID);
	if (size = DecompressBuffer((void *)ptr, buf, (DWORD)size, pwd))
		_o_return(BufferObject::Create(buf, size));
	_o_throw_oom;
}

BIF_DECL(BIF_ZipAddBuffer)
{
	size_t ptr, size;
	LPTSTR dest;
	if (auto obj = TokenToObject(*aParam[1]))
	{
		GetBufferObjectPtr(aResultToken, obj, (size_t&)ptr, size);
		if (aResultToken.Exited())
			return;
		dest = ParamIndexToOptionalString(2);
	}
	else {
		ptr = (size_t)TokenToInt64(*aParam[1]);
		size = (size_t)TokenToInt64(*aParam[2]);
		dest = ParamIndexToOptionalString(3);
	}
	if (ptr < 65536 || !size)
		_o_throw_value(ERR_INVALID_VALUE);
	auto hzip = (HZIP)TokenToInt64(*aParam[0]);
	if (auto aErrCode = ZipAddBuffer(hzip, dest, (const void *)ptr, (ULONGLONG)size))
		_o_throw_zip(aErrCode);
}

BIF_DECL(BIF_ZipAddFile)
{
	if (auto aErrCode = ZipAddFile((HZIP)TokenToInt64(*aParam[0]),
		ParamIndexIsOmittedOrEmpty(2) ? NULL : TokenToString(*aParam[2]),
		TokenToString(*aParam[1])))
		_o_throw_zip(aErrCode);
}

BIF_DECL(BIF_ZipAddFolder)
{
	if (auto aErrCode = ZipAddFolder((HZIP)TokenToInt64(*aParam[0]), TokenToString(*aParam[1])))
		_o_throw_zip(aErrCode);
}

BIF_DECL(BIF_ZipCloseBuffer)
{
	void		*aBuffer;
	ULONGLONG	aLen;
	HANDLE		aBase;
	if (auto aErrCode = ZipGetMemory((HZIP)TokenToInt64(*aParam[0]), &aBuffer, &aLen, &aBase))
		_o_throw_zip(aErrCode);
	LPVOID aData = malloc((size_t)aLen);
	if (aData)
		memcpy(aData, aBuffer, (size_t)aLen);
	// Free the memory now that we're done with it.
	UnmapViewOfFile(aBuffer);
	CloseHandle(aBase);
	if (aData)
		_o_return(BufferObject::Create(aData, (size_t)aLen));
	_o_throw_oom;
}

BIF_DECL(BIF_ZipCloseFile)
{
	if (auto aErrCode = ZipClose((HZIP)TokenToInt64(*aParam[0])))
		_o_throw_zip(aErrCode);
}

BIF_DECL(BIF_ZipCreateBuffer)
{
	HZIP hz;
	if (auto aErrCode = ZipCreateBuffer(&hz, 0, (ULONGLONG)TokenToInt64(*aParam[0]), CStringCharFromTChar(ParamIndexToOptionalString(1)), ParamIndexToOptionalInt(2, 5)))
		_o_throw_zip(aErrCode);
	_o_return((__int64)hz);
}

BIF_DECL(BIF_ZipCreateFile)
{
	HZIP hz;
	if (auto aErrCode = ZipCreateFile(&hz, TokenToString(*aParam[0]), CStringCharFromTChar(ParamIndexToOptionalString(1)), ParamIndexToOptionalInt(2, 5)))
		_o_throw_zip(aErrCode);
	_o_return((__int64)hz);
}

BIF_DECL(BIF_ZipInfo)
{
	IObject *obj;
	size_t	ptr = 0, size;
	DWORD	err;
	HZIP	huz;
	LPTSTR	path = nullptr;
	ExprTokenType aValue;
	if (obj = TokenToObject(*aParam[0]))
	{
		GetBufferObjectPtr(aResultToken, obj, ptr, size);
		if (aResultToken.Exited())
			return;
		if (ptr < 65536 || !size)
			_o_throw_value(ERR_PARAM1_INVALID);
	}
	else if (TokenIsPureNumeric(*aParam[0]))
	{
		if (!TokenIsNumeric(*aParam[1]))
			_o_throw_param(1, _T("Integer"));
		if (ptr = (size_t)TokenToInt64(*aParam[0]))
		{
			size = (size_t)TokenToInt64(*aParam[1]);
			if (ptr < 65536)
				_o_throw_value(ERR_PARAM1_INVALID);
			if (!size)
				_o_throw_value(ERR_PARAM2_INVALID);
		}
		aParam++;
		aParamCount--;
	}
	else
		path = TokenToString(*aParam[0]);
	auto codepage = (UINT)ParamIndexToOptionalInt(1, 0);
	if (path)
		err = UnzipOpenFileW(&huz, path, NULL, codepage);
	else
		err = UnzipOpenBuffer(&huz, (void *)ptr, (ULONGLONG)size, NULL, codepage); 
	if (err)
		goto error;

	ZIPENTRY	ze;
	ULONGLONG	numitems;
	Array *arr = Array::Create();
	auto buf = aResultToken.buf;
	ze.Index = (ULONGLONG)-1;
	if ((err = UnzipGetItem(huz, &ze)))
		goto errorclose;
	numitems = ze.Index;

	// Get info for all item(s).
	for (ze.Index = 0; ze.Index < numitems; ze.Index++)
	{
		if ((err = UnzipGetItem(huz, &ze)))
			goto errorclose;
		Object *obj = Object::Create();
		obj->SetOwnProp(_T("AccessTime"), FileTimeToYYYYMMDD(buf, ze.AccessTime, true));
		obj->SetOwnProp(_T("Attributes"), FileAttribToStr(buf, ze.Attributes));
		obj->SetOwnProp(_T("CompressedSize"), ze.CompressedSize);
		obj->SetOwnProp(_T("CreateTime"), FileTimeToYYYYMMDD(buf, ze.CreateTime, true));
		obj->SetOwnProp(_T("ModifyTime"), FileTimeToYYYYMMDD(buf, ze.ModifyTime, true));
		obj->SetOwnProp(_T("Name"), ze.Name);
		obj->SetOwnProp(_T("UncompressedSize"), ze.UncompressedSize);
		aValue.SetValue(obj);
		arr->Append(aValue);
		obj->Release();
	}

	// Done unzipping files, so close the ZIP archive.
	UnzipClose(huz);
	_o_return(arr);
errorclose:
	UnzipClose(huz);
	arr->Release();
error:
	_o_throw_zip(err);
}

BIF_DECL(BIF_ZipOptions)
{
	if (auto aErrCode = ZipOptions((HZIP)TokenToInt64(*aParam[0]), (DWORD)TokenToInt64(*aParam[1])))
		_o_throw_zip(aErrCode);
}

BIF_DECL(BIF_ZipRawMemory)
{
	size_t  ptr, size;
	LPVOID buf;
	LPTSTR pwd;
	int level;

	if (auto obj = TokenToObject(**aParam))
	{
		GetBufferObjectPtr(aResultToken, obj, ptr, size);
		if (aResultToken.Exited())
			return;
		pwd = ParamIndexIsOmittedOrEmpty(1) ? NULL : TokenToString(*aParam[1]);
		level = ParamIndexToOptionalInt(2, 5);
	}
	else
	{
		ptr = (size_t)TokenToInt64(*aParam[0]);
		size = (DWORD)TokenToInt64(*aParam[1]);
		pwd = ParamIndexIsOmittedOrEmpty(2) ? NULL : TokenToString(*aParam[2]);
		level = ParamIndexToOptionalInt(3, 5);
	}
	if (ptr < 65536 || !size)
		_o_throw_value(ERR_INVALID_VALUE);
	if (size = CompressBuffer((BYTE *)ptr, buf, (DWORD)size, pwd, level))
		_o_return(BufferObject::Create(buf, size));
	_o_throw_oom;
}
#undef _o_throw_zip



ResultType DynaToken::Invoke(IObject_Invoke_PARAMS_DECL)
{
	if (IS_INVOKE_CALL && !(aName && _tcsicmp(aName, _T("Call")))) {
		if (aParamCount > mParamCount)
			return aResultToken.Error(ERR_PARAM_COUNT_INVALID);
	}
	else return Object::Invoke(IObject_Invoke_PARAMS);

	aResultToken.SetValue(this);
	aResultToken.func = nullptr;
	BIF_DllCall(aResultToken, aParam, aParamCount);
	return aResultToken.Result();
}



thread_local IObject *JSON::_true;
thread_local IObject *JSON::_false;
thread_local IObject *JSON::_null;

void JSON::Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	JSON js;
	switch (aID)
	{
	case M_Parse:
		_o_return_retval(js.Parse(aResultToken, aParam, aParamCount));
	case M_Stringify:
		_o_return_retval(js.Stringify(aResultToken, aParam, aParamCount));
	case P_True:
		JSON::_true->AddRef();
		_o_return(JSON::_true);
	case P_False:
		JSON::_false->AddRef();
		_o_return(JSON::_false);
	case P_Null:
		JSON::_null->AddRef();
		_o_return(JSON::_null);
	}
}

void JSON::Parse(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	ExprTokenType val{}, _true(JSON::_true), _false(JSON::_false), _null(JSON::_null), _miss;
	size_t keylen, len, commas = 0;
	LPTSTR b, beg, key = NULL, t = NULL;
	Object::String valbuf, keybuf;
	TCHAR c, quot, endc, tp = '\0', nul[5] = {};
	bool success = true, escape = false, polarity, nonempty = false, as_map = ParamIndexToOptionalBOOL(2, TRUE);
	Array stack, *top = &stack;
	Object *temp;
	auto &deep = stack.mLength;
	union
	{
		IObject *cur;
		Object *cur_obj;
		Array *cur_arr;
		Map *cur_map;
	};
	cur = top;
	b = beg = TokenIsPureNumeric(*aParam[0]) ? (LPTSTR)TokenToInt64(*aParam[0]) : TokenToString(*aParam[0]);
	if (b < (void *)65536)
		_o_throw_value(ERR_PARAM1_INVALID);
	if (!ParamIndexToOptionalBOOL(1, TRUE))
		_true.SetValue(1), _false.SetValue(0), _null.SetValue(_T(""), 0);
	valbuf.SetCapacity(0x1000), keybuf.SetCapacity(0x1000);
	_miss.symbol = SYM_MISSING;

#define ltrim while ((c = *beg) == ' ' || c == '\t' || c == '\r' || c == '\n') beg++;

	auto trim_comment = [&]() {
		auto t = beg[1];
		if (t == '/')
			beg = _tcschr(beg += 2, '\n');
		else if (t == '*')
			beg = _tcsstr(beg += 2, _T("*/")) + 1;
		else return false;
		if (beg > (void *)65535)
			return true;
		beg = nul;
		return t == '/';
	};

	for (val.symbol = SYM_INVALID; ; beg++) {
		ltrim;
		switch (c)
		{
		case '/':
			if (!trim_comment())
				goto error;
			break;
		case '\0':
			if (deep != 1)
				goto error;
			stack.mItem[0].ReturnMove(aResultToken);
			goto ret;
		case '{':
		case '[':
			if (val.symbol != SYM_INVALID)
				goto error;
			if (tp == '}') {
				if (!key)
					goto error;
				temp = c == '[' ? Array::Create() :
					as_map ? Map::Create(nullptr, 0, true) :
					Object::Create(nullptr, 0, nullptr, true);
				success = as_map ? cur_map->SetItem(key, temp) : cur_obj->SetOwnProp(key, temp);
				key[keylen] = endc, key = NULL;
			}
			else {
				for (; commas; commas--)
					if (!(success = cur_arr->Append(_miss)))
						goto error;
				temp = c == '[' ? Array::Create() :
					as_map ? Map::Create(nullptr, 0, true) :
					Object::Create(nullptr, 0, nullptr, true);
				success = cur_arr->Append(ExprTokenType(temp));
			}
			temp->Release();
			if (!success)
				goto error;
			tp = c + 2, nonempty = false, commas = 0;
			stack.Append(ExprTokenType(cur = temp));
			break;
		case '}':
		case ']':
			if (c != tp)
				goto error;
			cur->Release();
			val.symbol = SYM_OBJECT, nonempty = true;
			stack.mItem[--deep].symbol = SYM_MISSING;
			if (deep == 1)
				tp = '\0', cur = top;
			else
				cur = stack.mItem[deep - 1].object, tp = cur_obj->mBase == Array::sPrototype ? ']' : '}';
			break;
		case '\'':
		case '"':
			if (val.symbol != SYM_INVALID)
				goto error;
			t = ++beg, escape = false, quot = c;
			while ((c = *beg) != quot) {
				if (c == '\\') {
					if (!(c = *(++beg)))
						goto error;
					escape = true;
				}
				else if (c == '\r' || c == '\n' && beg[-1] != '\r')
					goto error;
				beg++;
			}
			len = beg - t;
			if (escape) {
				auto &buf = tp != '}' || key ? valbuf : keybuf;
				if (len > buf.Capacity())
					buf.SetCapacity(max(len + 1, buf.Capacity() * 2));
				beg = t;
				LPTSTR p = t = buf;
				while ((c = *beg) != quot) {
					if (c == '\\') {
						switch (c = *(++beg))
						{
						case 'b': *p++ = '\b'; break;
						case 'f': *p++ = '\f'; break;
						case 'n': *p++ = '\n'; break;
						case 'r': *p++ = '\r'; break;
						case 't': *p++ = '\t'; break;
						case 'v': *p++ = '\v'; break;
						case 'u':
						case 'x':
							*p = 0;
							for (int i = c == 'u' ? 4 : 2; i; i--) {
								*p *= 16, c = *++beg;
								if (c >= '0' && c <= '9')
									*p += c - '0';
								else if (c >= 'A' && c <= 'F')
									*p += c - 'A' + 10;
								else if (c >= 'a' && c <= 'f')
									*p += c - 'a' + 10;
								else
									goto error;
							}
							p++;
							break;
						case '\r': beg[1] == '\n' && beg++;
						case '\n': break;
						case '0': c = 0;
						default: *p++ = c; break;
						}
					}
					else
						*p++ = c;
					beg++;
				}
				*p = '\0', len = p - t;
			}
			else
				t[len] = '\0';
			if (tp == '}') {
				if (!key) {
					key = t, keylen = len, endc = quot, beg++;
					while (true) {
						ltrim;
						if (c == ':')
							break;
						if (c == '/' && trim_comment() && beg++)
							continue;
						goto error;
					}
				}
				else {
					val.SetValue(t, len);
					success = as_map ? cur_map->SetItem(key, val) : cur_obj->SetOwnProp(key, val);
					key[keylen] = t[len] = endc, key = NULL;
				}
			}
			else {
				for (; commas; commas--)
					if (!(success = cur_arr->Append(_miss)))
						goto error;
				val.SetValue(t, len), success = cur_arr->Append(val), t[len] = quot;
			}
			if (!success)
				goto error;
			nonempty = true;
			break;
		case ',':
			if (cur == top)
				goto error;
			if (val.symbol == SYM_INVALID)
				commas++;
			else
				commas = 0, val.symbol = SYM_INVALID;
			break;
		default:
#define expect_str(str)               \
	for (int i = 0; str[i] != 0; ++i) \
	{                                 \
		if (str[i] != c)              \
			goto error;               \
		c = *++beg;                   \
	}
			if (val.symbol != SYM_INVALID)
				goto error;
			if (tp == '}' && !key) {
				auto end = find_identifier_end(beg);
				if (keylen = end - beg) {
					beg = end;
					while (true) {
						ltrim;
						if (c != ':')
							if (c == '/' && trim_comment() && beg++)
								continue;
							else
								goto error;
						break;
					}
					key = end - keylen;
					endc = *end, *end = '\0';
					nonempty = true;
					break;
				}
				goto error;
			}
			switch (c)
			{
			case 'f':
			case 'n':
			case 't':
				if (c == 'f') {
					expect_str(_T("false"));
					val.CopyValueFrom(_false);
				}
				else if (c == 'n') {
					expect_str(_T("null"));
					if (tp == ']')
						val.symbol = SYM_MISSING, cur_obj->SetOwnProp(_T("Default"), _null);
					else val.CopyValueFrom(_null);
				}
				else {
					expect_str(_T("true"));
					val.CopyValueFrom(_true);
				}
				nonempty = true, beg--;
				if (tp == '}') {
					success = as_map ? cur_map->SetItem(key, val) : cur_obj->SetOwnProp(key, val);
					key[keylen] = endc, key = NULL;
				}
				else {
					for (; commas; commas--)
						if (!(success = cur_arr->Append(_miss)))
							goto error;
					success = cur_arr->Append(val);
				}
				if (!success)
					goto error;
				break;
			default:
				t = beg, polarity = c == '-';
				if (polarity || c == '+')
					c = *++beg;
				if (c >= '0' && c <= '9') {
					bool is_hex = IS_HEX(beg);
					val.SetValue(istrtoi64(t, &beg));
					if (!is_hex) c = *beg;
				}
				else if (c != '.') {
					if (c == 'N') {
						expect_str(_T("NaN"));
						val.value_int64 = 0xffffffffffffffff;
					}
					else {
						expect_str(_T("Infinity"));
						val.value_int64 = polarity ? 0xfff0000000000000 : 0x7ff0000000000000;
					}
					val.symbol = SYM_FLOAT, c = 0;
				}
				if (c == '.' || c == 'e' || c == 'E')
					val.SetValue(_tcstod(t, &beg));
				beg--, nonempty = true;
				if (tp == '}') {
					success = as_map ? cur_map->SetItem(key, val) : cur_obj->SetOwnProp(key, val);
					key[keylen] = endc, key = NULL;
				}
				else {
					for (; commas; commas--)
						if (!(success = cur_arr->Append(_miss)))
							goto error;
					success = cur_arr->Append(val);
				}
				if (!success)
					goto error;
				break;
			}
		}
	}
error:
	if (!success)
		aResultToken.MemoryError();
	else {
		TCHAR msg[60];
		if (*beg) {
			if (*beg >= '0' && *beg <= '9')
				sntprintf(msg, 60, _T("Unexpected number in JSON at position %d"), beg - b);
			else
				sntprintf(msg, 60, _T("Unexpected token %c in JSON at position %d"), *beg, beg - b);
		}
		else
			_tcscpy(msg, _T("Unexpected end of JSON input"));
		aResultToken.ValueError(msg);
	}
ret:
	if (key) key[keylen] = endc;
}

void JSON::Stringify(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	if (IObject *obj = TokenToObject(*aParam[0])) {
		if (!ParamIndexIsOmitted(1)) {
			ExprTokenType val = *aParam[1];
			if (auto opts = dynamic_cast<Object *>(TokenToObject(val))) {
				if (opts->GetOwnProp(val, _T("depth")))
					depth = (UINT)TokenToInt64(val), indent = _T("\t"), indent_len = 1;
				if (!opts->GetOwnProp(val, _T("indent")))
					val.symbol = SYM_MISSING;
			}
			if (val.symbol != SYM_MISSING) {
				if (TokenIsNumeric(val)) {
					auto sz = TokenToInt64(val);
					if (sz > 0) {
						indent = (LPTSTR)_alloca(sizeof(TCHAR) * ((size_t)sz + 1));
						tmemset(indent, ' ', (size_t)sz);
						indent[indent_len = (UINT)sz] = '\0';
					}
				}
				else if (!TokenIsEmptyString(val))
					indent = TokenToString(val), indent_len = (UINT)_tcslen(indent);
			}
		}
		objcolon = indent ? _T("\": ") : _T("\":");
		deep = 0;
		appendObj(obj, ParamIndexToOptionalBOOL(2, true));
	}
	else if (TokenIsPureNumeric(*aParam[0])) {
		TCHAR numbuf[MAX_NUMBER_LENGTH];
		append(TokenToString(*aParam[0], numbuf));
	}
	else str.append('"'), append(TokenToString(*aParam[0])), str.append('"');
	ToToken(aResultToken);
}

void JSON::append(LPTSTR s) {
	static const char hexDigits[16] = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'a', 'b', 'c', 'd', 'e', 'f' };
	static const char escape[96] = {
		//0    1    2    3    4    5    6    7    8    9    A    B    C    D    E    F
		'u', 'u', 'u', 'u', 'u', 'u', 'u', 'u', 'b', 't', 'n', 'u', 'f', 'r', 'u', 'u', // 00
		'u', 'u', 'u', 'u', 'u', 'u', 'u', 'u', 'u', 'u', 'u', 'u', 'u', 'u', 'u', 'u', // 10
		  0,   0, '"',   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0, // 20
		  0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0, // 30
		  0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0, // 40
		  0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,   0,'\\',   0,   0,   0, // 50~5F
	};
	TCHAR escbuf[7] = _T("\\u0000");
	LPTSTR b;
	TCHAR c, ch;
	size_t sz;

	for (b = s; ch = *s; s++) {
		if (ch < 0x60 && (c = escape[ch])) {
			if (b) {
				if (sz = s - b)
					str.append(b, sz);
				b = NULL;
			}
			if (c == 'u') {
				escbuf[4] = hexDigits[ch >> 4];
				escbuf[5] = hexDigits[ch & 0xf];
				str += escbuf;
			}
			else
				str.append('\\'), str.append(c);
		}
		else if (!b)
			b = s;
	}
	if (b && (sz = s - b))
		str.append(b, sz);
}

const IID IID_IObjectComCompatible;
void JSON::appendObj(IObject *obj, bool invoke_tojson) {
	if (dynamic_cast<EnumBase *>(obj))
		return append(obj, true);
	ResultToken result;
	if (auto m = dynamic_cast<Map *>(obj))
		append(m);
	else if (auto a = dynamic_cast<Array *>(obj))
		append(a);
	else if (auto o = dynamic_cast<Object *>(obj))
		o->HasMethod(_T("__Enum")) ? append(obj) : append(o);
	else {
		auto comobj = dynamic_cast<ComObject *>(obj);
		if (!comobj || (comobj->mVarType == VT_NULL || comobj->mVarType == VT_EMPTY))
			str += _T("null");
		else if (comobj->mVarType == VT_BOOL)
			str += comobj->mVal64 == VARIANT_FALSE ? _T("false") : _T("true");
		else if (comobj->mVarType == VT_DISPATCH || (comobj->mVarType & VT_ARRAY)) {
			if (invoke_tojson) {
				if (comobj->mVarType == VT_DISPATCH)
					if (SUCCEEDED(comobj->mDispatch->QueryInterface(IID_IObjectComCompatible, (void **)&obj)))
						obj->Release();
					else invoke_tojson = false;
				if (invoke_tojson) {
					ExprTokenType t_this((__int64)this);
					ExprTokenType *param[] = { &t_this };
					auto prev_excpt = g->ExcptMode;
					auto l = str.size();
					g->ExcptMode = EXCPTMODE_CATCH;
					t_this.symbol = SYM_PTR;
					result.InitResult(_T(""));
					if (obj->Invoke(result, IT_CALL, _T("ToJSON"), ExprTokenType(obj), param, 1) == OK && result.Result() == OK && l < str.size()) {
						result.Free();
						g->ExcptMode = prev_excpt;
						return;
					}
					result.Free();
					g->ExcptMode = prev_excpt;
					if (g->ThrownToken)
						g_script->FreeExceptionToken(g->ThrownToken);
				}
			}
			append(obj);
		}
		else
			str += _T("null");
	}
}

bool JSON::append(Variant &field, Object *obj) {
	TCHAR numbuf[MAX_NUMBER_LENGTH];
	switch (field.symbol)
	{
	case SYM_OBJECT:
		appendObj(field.object);
		break;
	case SYM_STRING:
		str.append('"');
		append(field.string);
		str.append('"');
		break;
	case SYM_FLOAT:
		FTOA(field.n_double, numbuf, MAX_NUMBER_LENGTH);
		str += numbuf;
		break;
	case SYM_INTEGER:
		ITOA64(field.n_int64, numbuf);
		str += numbuf;
		break;
	case SYM_DYNAMIC:
		if (obj && field.prop->Getter() && !field.prop->NoEnumGet && !obj->IsClassPrototype()) {
			ResultToken result_token;
			ExprTokenType getter(field.prop->Getter());
			ExprTokenType object(obj);
			auto param = &object;
			result_token.InitResult(numbuf);
			auto result = getter.object->Invoke(result_token, IT_CALL, nullptr, getter, &param, 1);
			if (!(result == FAIL || result == EARLY_EXIT)) {
				switch (result_token.symbol)
				{
				case SYM_OBJECT:
					appendObj(result_token.object);
					result_token.object->Release();
					return true;
				case SYM_STRING:
					str.append(result_token.marker, result_token.marker_length);
					free(result_token.mem_to_free), result_token.mem_to_free = nullptr;
					return true;
				case SYM_FLOAT:
					FTOA(result_token.value_double, numbuf, MAX_NUMBER_LENGTH);
					str += numbuf;
					return true;
				case SYM_INTEGER:
					ITOA64(result_token.value_int64, numbuf);
					str += numbuf;
					return true;
				}
			}
		}
	default: return false;
	}
	return true;
}

void JSON::append(Object *obj) {
	auto &mFields = obj->mFields;
	str.append('{'), ++deep;
	for (index_t i = 0; i < mFields.Length(); i++) {
		FieldType &field = mFields[i];
		auto sz = str.size();
		append(deep), str.append('"');
		append(field.name);
		str += objcolon;
		if (append(field, obj))
			str.append(',');
		else str.size() = sz;
	}
	--deep;
	if (str.back() == ',')
		--str.size(), append(deep);
	str.append('}');
}

void JSON::append(Array *obj) {
	auto &mItem = obj->mItem;
	auto mLength = obj->Length();
	str.append('['), ++deep;
	for (index_t i = 0; i < mLength; i++)
		append(deep), append(mItem[i]), str.append(',');
	--deep;
	if (str.back()  == ',')
		--str.size(), append(deep);
	str.append(']');
}

void JSON::append(Map *obj) {
	TCHAR numbuf[MAX_NUMBER_LENGTH];
	auto mItem = obj->mItem;
	auto mCount = obj->mCount;
	auto IsUnsorted = obj->IsUnsorted();
	str.append('{'), ++deep;
	for (index_t i = 0; i < mCount; i++) {
		Map::Pair &pair = mItem[i];
		append(deep), str.append('"');
		switch (IsUnsorted ? (SymbolType)obj->mKeyTypes[i] : i < obj->mKeyOffsetObject ? SYM_INTEGER : i < obj->mKeyOffsetString ? SYM_OBJECT : SYM_STRING)
		{
		case SYM_STRING:
			append(pair.key.s);
			break;
		case SYM_OBJECT:
			str += pair.key.p->Type();
			str.append('_');
		case SYM_INTEGER:
			ITOA64(pair.key.i, numbuf);
			str += numbuf;
			break;
		}
		str += objcolon;
		append(pair);
		str.append(',');
	}
	--deep;
	if (str.back() == ',')
		--str.size(), append(deep);
	str.append('}');
}

void JSON::append(Var &vval) {
	append(deep);
	if (vval.IsObject())
		appendObj(vval.mObject);
	else if (vval.IsPureNumeric())
		append(vval.Contents());
	else
		str.append('"'), append(vval.Contents()), str.append('"');
	str.append(',');
}

void JSON::append(Var &vkey, Var &vval) {
	append(deep), str.append('"');
	if (vkey.IsObject()) {
		TCHAR numbuf[MAX_NUMBER_SIZE];
		str += vkey.mObject->Type(), str.append('_');
		ITOA64((__int64)vkey.ToObject(), numbuf);
		str += numbuf;
	}
	else
		append(vkey.Contents());
	str += objcolon;
	if (vval.IsObject())
		appendObj(vval.mObject);
	else if (vval.IsPureNumeric())
		append(vval.Contents());
	else
		str.append('"'), append(vval.Contents()), str.append('"');
	str.append(',');
}

void JSON::append(IObject *obj, bool is_enum_base) {
	IObject *enumerator = NULL;
	ResultToken result_token;
	ExprTokenType params[3], enum_token(obj);
	ExprTokenType *param[] = { params,params + 1,params + 2 };
	TCHAR buf[MAX_NUMBER_LENGTH];
	result_token.InitResult(buf);
	ResultType result;
	Var vkey, vval;
	auto prev_excpt = g->ExcptMode;

	g->ExcptMode = EXCPTMODE_CATCH;
	if (is_enum_base)
		enumerator = obj, obj->AddRef();
	if (enumerator || (obj && GetEnumerator(enumerator, enum_token, 2, false) == OK)) {
		IndexEnumerator *enumobj = dynamic_cast<IndexEnumerator *>(enumerator);
		if (enumobj || dynamic_cast<EnumBase *>(enumerator)) {
			char enumtype = enumobj ? 1 : 0;
			if (enumobj && dynamic_cast<Array *>(enumobj->mObject))
				enumobj->mParamCount = 2, enumtype++;
			if (enumtype == 1) {
				str.append('{'), ++deep;
				while (((EnumBase *)enumerator)->Next(&vkey, &vval) == CONDITION_TRUE)
					append(vkey, vval);
				vkey.Free(), vval.Free();
				if (str.back() == ',')
					str.pop_back(), append(--deep), str.append('}');
				else
					--deep, str.append('}');
			}
			else if (enumtype == 2 || dynamic_cast<ComArrayEnum *>(enumerator)) {
				if (!enumobj)
					((ComArrayEnum *)enumerator)->mIndexMode = true;
				str.append('['), ++deep;
				while (((EnumBase *)enumerator)->Next(&vkey, &vval) == CONDITION_TRUE)
					append(vval);
				vkey.Free(), vval.Free();
				if (str.back() == ',')
					str.pop_back(), append(--deep), str.append(']');
				else
					--deep, str.append(']');
			}
			else {
				auto e = '}';
				++deep;
				if (enumerator == obj)
					str.append('['), e = ']';
				else str.append('{');
				if (((ComEnum *)enumerator)->Next(&vkey, nullptr) == CONDITION_TRUE) {
					if (e == '}') {
						vkey.ToTokenSkipAddRef(params[0]);
						if (obj->Invoke(result_token, IT_GET, nullptr, enum_token, param, 1) == OK) {
							vval.Assign(result_token);
							append(vkey, vval);
							result_token.Free();
						}
						else {
							e = ']', str.back() = '[', append(vkey);
							if (g->ThrownToken)
								g_script->FreeExceptionToken(g->ThrownToken);
						}
					}
					else
						append(vkey);
					if (e == '}') {
						while (((ComEnum *)enumerator)->Next(&vkey, nullptr) == CONDITION_TRUE) {
							vkey.ToTokenSkipAddRef(params[0]);
							if (obj->Invoke(result_token, IT_GET, nullptr, enum_token, param, 1) == OK)
								vval.Assign(result_token), result_token.Free();
							else {
								vval.Assign(_T(""), 0);
								if (g->ThrownToken)
									g_script->FreeExceptionToken(g->ThrownToken);
							}
							append(vkey, vval);
						}
					}
					else {
						while (((ComEnum *)enumerator)->Next(&vkey, nullptr) == CONDITION_TRUE)
							append(vkey);
					}
					vkey.Free(), vval.Free();
				}
				if (str.back() == ',')
					str.pop_back(), append(--deep), str.append(e);
				else
					--deep, str.append(e);
			}
		}
		else {
			VarRef vkey, vval;
			ExprTokenType t_this(enumerator);
			params[0].SetValue(&vkey), params[1].SetValue(&vval);
			result_token.SetValue(0);
			result = OK;
			if (enumerator->Invoke(result_token, IT_CALL, nullptr, t_this, param, 2) == OK) {
				str.append('{'), ++deep;
				while (TokenToBOOL(result_token)) {
					append(vkey, vval);
					if (enumerator->Invoke(result_token, IT_CALL, nullptr, t_this, param, 2) != OK)
						break;
				}
				if (str.back() == ',')
					str.pop_back(), append(--deep), str.append('}');
				else
					--deep, str.append('}');
			}
			else
				str += _T("null");
			if (g->ThrownToken)
				g_script->FreeExceptionToken(g->ThrownToken);
			result_token.Free();
		}
		enumerator->Release();
	}
	else if (obj) {
		str += _T("null");
		if (g->ThrownToken)
			g_script->FreeExceptionToken(g->ThrownToken);
	}
	g->ExcptMode = prev_excpt;
}



bool Promise::Init(ExprTokenType *aParam[], int aParamCount) {
	if (!aParamCount)
		return true;
	mParamToken = (ResultToken *)malloc(aParamCount * (sizeof(ResultToken *) + sizeof(ResultToken)));
	if (!mParamToken)
		return false;
	mParam = (ResultToken **)(mParamToken + aParamCount);
	for (auto &i = mParamCount; i < aParamCount; i++) {
		auto &token = mParamToken[i];
		mParam[i] = &token;
		token.CopyExprFrom(*aParam[i]);
		if (token.symbol == SYM_VAR)
			token.var->ToTokenSkipAddRef(token);
		if (token.symbol == SYM_STRING)
			token.Malloc(token.marker, token.marker_length);
		else {
			token.mem_to_free = nullptr;
			if (token.symbol == SYM_OBJECT) {
				if (!mMarshal)
					token.object->AddRef();
				else if (!MarshalObjectToToken(token.object, token))
					return false;
			}
		}
	}
	return true;
}

void Promise::Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	Object *callback = dynamic_cast<Object *>(TokenToObject(*aParam[0]));
	if (!callback)
		_o_throw(ERR_INVALID_FUNCTOR);
	if (!ValidateFunctor(callback, 1, aResultToken))
		return;
	EnterCriticalSection(&mCritical);
	auto &stream_to_set = mStream[aID - 1];
	auto &callback_to_set = aID == M_Then ? mComplete : mError;
	IStream *pstm = nullptr;
	callback_to_set = callback;
	if (!(mState & aID)) {
		if (FAILED(CoMarshalInterThreadInterfaceInStream(IID_IDispatch, callback, &pstm))) {
			callback_to_set = nullptr;
			LeaveCriticalSection(&mCritical);
			_o_throw_oom;
		}
		callback = nullptr;
	}
	if (stream_to_set) {
		CoReleaseMarshalData(stream_to_set);
		stream_to_set->Release();
	}
	stream_to_set = pstm;
	LeaveCriticalSection(&mCritical);
	if (callback) {
		AddRef();
		mReply = (HWND)-1;
		OnFinish();
	}
	AddRef();
	_o_return(this);
}

void Promise::FreeParams() {
	LPSTREAM stream;
	for (int i = 0; i < mParamCount; i++) {
		if (mParam[i]->symbol == SYM_STREAM) {
			stream = (LPSTREAM)mParam[i]->value_int64;
			CoReleaseMarshalData(stream);
			stream->Release();
		}
		else mParam[i]->Free();
	}
	mParamCount = 0;
}

void Promise::FreeResult() {
	LPSTREAM stream;
	if (mResult.symbol == SYM_STREAM) {
		stream = (LPSTREAM)mResult.value_int64;
		CoReleaseMarshalData(stream);
		stream->Release();
	}
	else mResult.Free();
	mResult.InitResult(_T(""));
}

void Promise::UnMarshal() {
	if (!mMarshal)
		return;
	for (int i = 0; i < mParamCount; i++)
		if (mParam[i]->symbol == SYM_STREAM)
			if (auto pobj = UnMarshalObjectFromStream((LPSTREAM)mParam[i]->value_int64))
				mParam[i]->SetValue(pobj);
			else mParam[i]->SetValue(_T(""));
}

static BuiltInFunc sPromiseCaller(_T(""), [](BIF_DECL_PARAMS) {
	auto p = (Promise *)aParam[0]->object;
	if (p->mState)
		p->OnFinish();
	else p->Call();
	}, 1, 1);
Func *Promise::ToBoundFunc() {
	ExprTokenType t_this(this), *param[] = { &t_this };
	return BoundFunc::Bind(&sPromiseCaller, IT_CALL, nullptr, param, 1);
}

void callFuncDll(FuncAndToken &aToken, bool throwerr);
void Promise::Call()
{
	UnMarshal();
	FuncAndToken token = { mFunc,&mResult,(ExprTokenType **)mParam,mParamCount };
	callFuncDll(token, false);
	FreeParams();
	EnterCriticalSection(&mCritical);
	mState = mResult.Result() == FAIL ? 2 : 1;
	mFunc->Release();
	mFunc = nullptr;
	if (auto hwnd = mReply) {
		if (mMarshal && mResult.symbol == SYM_OBJECT) {
			auto obj = mResult.object;
			if (!MarshalObjectToToken(obj, mResult))
				mResult.SetValue(_T(""));
			obj->Release();
		}
		else if (mResult.symbol == SYM_STRING && !mResult.mem_to_free)
			mResult.Malloc(mResult.marker, mResult.marker_length);
		mReply = (HWND)-1;
		LeaveCriticalSection(&mCritical);
		DWORD_PTR res;
		if (SendMessageTimeout(hwnd, AHK_EXECUTE_FUNCTION, 0, (LPARAM)this, SMTO_ABORTIFHUNG, 5000, &res))
			return;
	}
	else
		LeaveCriticalSection(&mCritical);
	FreeResult();
	Release();
}

void Promise::OnFinish()
{
	if (mResult.symbol == SYM_STREAM) {
		if (auto pobj = UnMarshalObjectFromStream((LPSTREAM)mResult.value_int64))
			mResult.SetValue(pobj);
		else mResult.SetValue(_T(""));
	}
	if (IObject *callback = (mState & 2) ? mError : mComplete) {
		ExprTokenType *param[] = { &mResult };
		ResultToken result;
		TCHAR buf[MAX_INTEGER_LENGTH];
		mState |= 4;
		result.InitResult(buf);
		callback->Invoke(result, IT_CALL, nullptr, ExprTokenType(callback), param, 1);
		result.Free();
	}
	Release();
}

Worker *Worker::Create()
{
	auto obj = new Worker();
	obj->SetBase(sPrototype);
	return obj;
}

bool Worker::New(DWORD aThreadID)
{
	mIndex = 0;
	if (aThreadID)
		for (; mIndex < MAX_AHK_THREADS && g_ahkThreads[mIndex].ThreadID != aThreadID; ++mIndex);
	if (mIndex >= MAX_AHK_THREADS || !(mThreadID = g_ahkThreads[mIndex].ThreadID) || !(mHwnd = g_ahkThreads[mIndex].Hwnd))
		return false;
	return !!(mThread = OpenThread(THREAD_ALL_ACCESS, FALSE, mThreadID));
}

bool Worker::New(ResultToken &aResultToken, LPTSTR aScript, LPTSTR aCmd, LPTSTR aTitle)
{
	IObjPtr err{};
	void *param[] = { g,ErrorPrototype::Error };
	if (DWORD threadid = NewThread(aScript, aCmd, aTitle, TRUE,
		[](void *param) {
			if (g->ThrownToken) {
				auto obj = (Object *)g->ThrownToken->object;
				obj->SetBase(nullptr);
				*(Object **)param = obj;
				g->ThrownToken->symbol = SYM_INTEGER;
				g_script->FreeExceptionToken(g->ThrownToken);
			}
		}, &err)) {
		return Worker::New(threadid);
	}
	if (auto obj = (Object *)err.obj) {
		ExprTokenType msg, extra;
		msg.SetValue(_T(""));
		extra.SetValue(_T(""));
		obj->GetOwnProp(msg, _T("Message"));
		obj->GetOwnProp(extra, _T("Extra"));
		aResultToken.Error(msg.marker, extra.marker);
	}
	return false;
}

ResultType Worker::GetEnumItem(UINT &aIndex, Var *aVal, Var *aReserved, int aVarCount)
{
	for (; aIndex < MAX_AHK_THREADS; ++aIndex)
	{
		if (!g_ahkThreads[aIndex].Hwnd || !g_ahkThreads[aIndex].ThreadID)
			continue;
		if (aVarCount > 1)
		{
			if (aVal)
				aVal->Assign((__int64)g_ahkThreads[aIndex].ThreadID);
			aVal = aReserved;
		}
		if (aVal)
		{
			auto obj = Worker::Create();
			if (obj->New(g_ahkThreads[aIndex].ThreadID))
			{
				aVal->Assign(obj);
				obj->Release();
			}
			else
			{
				obj->Release();
				continue;
			}
		}
		return CONDITION_TRUE;
	}
	return CONDITION_FALSE;
}

void Worker::Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	Script *script;
	DWORD ThreadID = GetCurrentThreadId();
	if (aID == M___New) {
		if (!aParamCount || ParamIndexIsNumeric(0)) {
			if (New((DWORD)ParamIndexToOptionalInt64(0, ThreadID)))
				_o_return_retval;
			_o_throw(_T("Invalid ThreadID"));
		}
		New(aResultToken, ParamIndexToString(0), ParamIndexToOptionalString(1), ParamIndexToOptionalStringDef(2, _T("AutoHotkey")));
		_o_return_retval;
	}
	if (mThreadID == ThreadID)
		script = g_script;
	else {
		if (!mHwnd || !mThread || g_ahkThreads[mIndex].Hwnd != mHwnd) {
			mHwnd = NULL;
			if (mThread && WaitForSingleObject(mThread, 50) == WAIT_TIMEOUT) {
				for (mIndex = 0; mIndex < MAX_AHK_THREADS && g_ahkThreads[mIndex].ThreadID != mThreadID; ++mIndex);
				if (mIndex < MAX_AHK_THREADS)
					mHwnd = g_ahkThreads[mIndex].Hwnd;
			}
			if (!mHwnd) {
				if (aID == M_ExitApp)
					_o_return_retval;
				if (aID == M_Wait || aID == P_Ready)
					_o_return(aID == M_Wait);
				_o_throw_win32(RPC_E_SERVER_DIED_DNE);
			}
		}
		script = g_ahkThreads[mIndex].Script;
		if (!script || !script->mIsReadyToExecute) {
			if (aID == P_Ready)
				_o_return(0);
			if (aID != M_Wait)
				_o_throw_win32(RPC_E_SERVER_DIED_DNE);
		}
	}

	ExprTokenType *val = nullptr;
	IObject *obj = nullptr;
	Var *var = nullptr;
	LPTSTR name;

	switch (aID)
	{
	case P___Item:
	{
		if (IS_INVOKE_SET)
			val = aParam[0], aParam++, aParamCount--;
		var = script->FindGlobalVar2(name = TokenToString(*aParam[0]));
		DWORD_PTR res;
		DWORD err = 0;

		if (!var)
			_o_throw(ERR_DYNAMIC_NOT_FOUND, name);
		if (IS_INVOKE_SET) {
			if (mThreadID != ThreadID) {
				ResultToken result{};
				if (obj = TokenToObject(*val)) {
					if (!MarshalObjectToToken(obj, result))
						_o_throw_win32();
					val = &result;
				}
				else result.symbol = SYM_INTEGER;
				if (!SendMessageTimeout(mHwnd, AHK_THREADVAR, (WPARAM)val, (LPARAM)var, SMTO_ABORTIFHUNG, 5000, &res))
					err = GetLastError();
				if (result.symbol == SYM_STREAM) {
					auto stream = (LPSTREAM)result.value_int64;
					CoReleaseMarshalData(stream);
					stream->Release();
				}
				else free(result.mem_to_free);
				if (err)
					_o_throw_win32(err);
				if (!res)
					_o_throw_oom;
			}
			else
				var->Assign(*val);
			_o_return_retval;
		}

		// __item.get
		if (mThreadID == ThreadID)
			var->Get(aResultToken);
		else {
			var = var->ResolveAlias();
			if (var->mType == VAR_VIRTUAL)
			{
				AutoTLS tls;
				tls.Enter(g_ahkThreads[mIndex].TLS);
				var->Get(aResultToken);
				if (aResultToken.symbol == SYM_OBJECT)
					aResultToken.object->Release(), aResultToken.symbol = SYM_INVALID;
			}
			else
			{
				aResultToken.symbol = SYM_INVALID;
				switch (var->mAttrib & VAR_ATTRIB_TYPES)
				{
				case VAR_ATTRIB_IS_INT64:
					aResultToken.SetValue(var->mContentsInt64);
					break;
				case VAR_ATTRIB_IS_DOUBLE:
					aResultToken.SetValue(var->mContentsDouble);
					break;
				default:
					if (var->mAttrib & VAR_ATTRIB_UNINITIALIZED)
						aResultToken.symbol = SYM_MISSING;
				}
			}
			if (aResultToken.symbol == SYM_INVALID) {
				if (!SendMessageTimeout(mHwnd, AHK_THREADVAR, (WPARAM)&aResultToken, (LPARAM)var, SMTO_ABORTIFHUNG, 5000, &res))
					_o_throw_win32();
				if (aResultToken.symbol == SYM_STREAM) {
					if (!(obj = UnMarshalObjectFromStream((LPSTREAM)aResultToken.value_int64)))
						_o_throw_win32();
					aResultToken.SetValue(obj);
				}
				else if (!res)
					_o_throw_oom;
			}
		}
		_o_return_retval;
	}

	case M_AddScript:
		_o_return(addScript(ParamIndexToString(0), ParamIndexToOptionalInt(1, 0), mThreadID));

	case M_AsyncCall:
	{
		auto func = mThreadID == ThreadID ? TokenToObject(*aParam[0]) : nullptr;
		if (!func) {
			var = script->FindGlobalVar2(name = TokenToString(*aParam[0]));
			if (!var)
				_o_throw(ERR_DYNAMIC_NOT_FOUND, name);
			if (mThreadID == ThreadID && !(func = dynamic_cast<Object *>(var->ToObject())))
				_o_throw(ERR_INVALID_FUNCTOR);
		}
		auto promise = new Promise(func, var, mThreadID != ThreadID);
		promise->mPriority = g->Priority;
		promise->mReply = g_hWnd;
		if (!promise->Init(++aParam, --aParamCount)) {
			promise->Release();
			_o_throw_oom;
		}

		DWORD_PTR res;
		if (!SendMessageTimeout(mHwnd, AHK_EXECUTE_FUNCTION, 0, (LPARAM)promise, SMTO_ABORTIFHUNG, 5000, &res)) {
			auto err = GetLastError();
			promise->Release();
			_o_throw_win32(err);
		}
		if (!res) {
			promise->Release();
			_o_throw(ERR_INVALID_FUNCTOR);
		}
		promise->AddRef();
		_o_return(promise);
	}

	case M_Exec:
		_o_return(ahkExec(ParamIndexToString(0), mThreadID != ThreadID ? mThreadID : ParamIndexToOptionalBOOL(1, TRUE) ? 0 : mThreadID));

	case M_ExitApp:
		if (script)
			TerminateSubThread(mThreadID, mHwnd);
		CloseHandle(mThread), mThread = NULL;
		_o_return_retval;

	case M_Pause:
		_o_return(ahkPause(ParamIndexToBOOL(0), mThreadID));

	case M_Reload:
		_o_return(PostMessage(mHwnd, WM_COMMAND, ID_FILE_RELOADSCRIPT, 0));

	case P_Ready:
		_o_return(script->mIsReadyToExecute);

	case P_ThreadID:
		_o_return(mThreadID);

	case M_Wait:
		if (mThreadID == ThreadID)
			_o_return(0);
		auto timeout = ParamIndexToOptionalInt64(0, 0);
		auto g = ::g;
		for (int start_time = GetTickCount(); WaitForSingleObject(mThread, 0) == WAIT_TIMEOUT;) {
			if (g->Exited())
				return (void)aResultToken.SetExitResult(EARLY_EXIT);
			if (timeout && (int)(timeout - (GetTickCount() - start_time)) <= SLEEP_INTERVAL_HALF)
				_o_return(0);
			MsgSleep(INTERVAL_UNSPECIFIED);
		}
		_o_return(1);
	}
}

bool MarshalObjectToToken(IObject *aObj, ResultToken &aToken) {
	bool idisp = true;
	IUnknown *marsha = aObj;
	LPSTREAM stream;
	aToken.mem_to_free = nullptr;
	if (auto comobj = dynamic_cast<ComObject *>(aObj)) {
		if (comobj->mVarType == VT_DISPATCH || comobj->mVarType == VT_UNKNOWN)
			marsha = comobj->mUnknown, idisp = comobj->mVarType == VT_DISPATCH;
		else if (!(comobj->mVarType & ~VT_TYPEMASK)) {
			VARIANT vv{};
			comobj->ToVariant(vv);
			VariantToToken(vv, aToken);
			marsha = nullptr;
		}
	}
	if (!marsha)
		return true;
	if (SUCCEEDED(CoMarshalInterThreadInterfaceInStream(idisp ? IID_IDispatch : IID_IUnknown, marsha, &stream)) ||
		(idisp && SUCCEEDED(CoMarshalInterThreadInterfaceInStream(IID_IUnknown, marsha, &stream))))
		aToken.value_int64 = (__int64)stream, aToken.symbol = SYM_STREAM;
	else return false;
	return true;
}