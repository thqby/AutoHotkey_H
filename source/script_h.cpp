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

BIF_DECL(BIF_Cast)
{
	__int64 aValue = 0;
	ExprTokenType& token_to_write = *aParam[1];

	size_t target = (size_t)&aValue; // Don't make target a pointer-type because the integer offset might not be a multiple of 4 (i.e. the below increments "target" directly by "offset" and we don't want that to use pointer math).

	size_t source_size = sizeof(DWORD_PTR); // Set defaults.
	int source_type = 0; // 1 Char; 2 Uchar; 3 Short; 4 UShort; 5 Int; 6 UInt; 7 Float; 8 Double; 9 Int64
	size_t target_size = sizeof(DWORD_PTR); // Set defaults.
	BOOL is_integer = TRUE;   //
	BOOL is_unsigned = FALSE;

	LPTSTR type = TokenToString(*aParam[0]); // No need to pass aBuf since any numeric value would not be recognized anyway.
	if (ctoupper(*type) == 'U') // Unsigned; but in the case of NumPut, it matters only for UInt64.
	{
		is_unsigned = TRUE;
		++type; // Remove the first character from further consideration.
	}

	switch (ctoupper(*type)) // Override "size" and is_integer if type warrants it. Note that the above has omitted the leading "U", if present, leaving type as "Int" vs. "Uint", etc.
	{
	case 'P':
		is_unsigned = TRUE;
#ifdef _WIN64
		source_type = 9;
#else
		source_type = 6;
#endif
		break; // Ptr.
	case 'I':
		if (_tcschr(type, '6')) // Int64. It's checked this way for performance, and to avoid access violation if string is bogus and too short such as "i64".
		{
			source_size = 8;
			source_type = 9;
		}
		else
		{
			source_size = 4;
			source_type = 5 + is_unsigned;
		}
		//else keep "size" at its default set earlier.
		break;
	case 'S': source_size = 2; source_type = 3 + is_unsigned; break; // Short.
	case 'C': source_size = 1; source_type = 1 + is_unsigned; break; // Char.

	case 'D': source_size = 8; is_integer = FALSE; source_type = 8;  break; // Double.
	case 'F': source_size = 4; is_integer = FALSE; source_type = 7; break; // Float.

		// default: For any unrecognized values, keep "size" and is_integer at their defaults set earlier
		// (for simplicity).
	}

	switch (source_size)
	{
	case 4: // Listed first for performance.
		if (is_integer)
			*(unsigned int*)target = (unsigned int)TokenToInt64(token_to_write);
		else // Float (32-bit).
			*(float*)target = (float)TokenToDouble(token_to_write);
		break;
	case 8:
		if (is_integer)
			// v1.0.48: Support unsigned 64-bit integers like DllCall does:
			*(__int64*)target = (is_unsigned && !IS_NUMERIC(token_to_write.symbol)) // Must not be numeric because those are already signed values, so should be written out as signed so that whoever uses them can interpret negatives as large unsigned values.
			? (__int64)ATOU64(TokenToString(token_to_write)) // For comments, search for ATOU64 in BIF_DllCall().
			: TokenToInt64(token_to_write);
		else // Double (64-bit).
			*(double*)target = TokenToDouble(token_to_write);
		break;
	case 2:
		if (is_unsigned)
			*(unsigned short*)target = (unsigned short)TokenToInt64(token_to_write);
		else
			*(short*)target = (short)TokenToInt64(token_to_write);
		break;
	default: // size 1
		if (is_unsigned)
			*(unsigned char*)target = (unsigned char)TokenToInt64(token_to_write);
		else
			*(char*)target = (char)TokenToInt64(token_to_write);
	}

	type = TokenToString(*aParam[2]); // No need to pass aBuf since any numeric value would not be recognized anyway.
	if (ctoupper(*type) == 'U') // Unsigned.
	{
		++type; // Remove the first character from further consideration.
		is_unsigned = TRUE;
	}
	else
		is_unsigned = FALSE;

	switch (ctoupper(*type)) // Override "size" and aResultToken.symbol if type warrants it. Note that the above has omitted the leading "U", if present, leaving type as "Int" vs. "Uint", etc.
	{
		//case 'P': // Nothing extra needed in this case.
	case 'I':
		if (_tcschr(type, '6')) // Int64. It's checked this way for performance, and to avoid access violation if string is bogus and too short such as "i64".
			target_size = 8;
		else
			target_size = 4;
		break;
	case 'S': target_size = 2; break; // Short.
	case 'C': target_size = 1; break; // Char.

	case 'D': target_size = 8; aResultToken.symbol = SYM_FLOAT; break; // Double.
	case 'F': target_size = 4; aResultToken.symbol = SYM_FLOAT; break; // Float.

		// default: For any unrecognized values, keep "size" and aResultToken.symbol at their defaults set earlier
		// (for simplicity).
	}
#define CONVERT_SOURCE (\
		source_type == 1 ? *(char *)target : source_type == 2 ? *(unsigned char *)target : source_type == 3 ? *(short *)target : source_type == 4 ? *(unsigned short *)target : source_type == 5 ? *(int *)target : source_type == 6 ? *(unsigned int *)target : source_type == 7 ? *(float *)target : source_type == 8 ? *(double *)target : *(__int64 *)target \
	)
	switch (target_size)
	{
	case 4: // Listed first for performance.
		if (aResultToken.symbol == SYM_FLOAT)
			aResultToken.value_double = (float)CONVERT_SOURCE;
		else if (!is_unsigned)
			aResultToken.value_int64 = (int)CONVERT_SOURCE; // aResultToken.symbol was set to SYM_FLOAT or SYM_INTEGER higher above.
		else
			aResultToken.value_int64 = (unsigned int)CONVERT_SOURCE;
		break;
	case 8:
		if (ctoupper(*type) == 'D')
			aResultToken.value_double = (double)CONVERT_SOURCE;
		else
			aResultToken.value_int64 = (__int64)CONVERT_SOURCE;
		break;
	case 2:
		if (!is_unsigned) // Don't use ternary because that messes up type-casting.
			aResultToken.value_int64 = (short)CONVERT_SOURCE;
		else
			aResultToken.value_int64 = (unsigned short)CONVERT_SOURCE;
		break;
	default: // size 1
		if (!is_unsigned) // Don't use ternary because that messes up type-casting.
			aResultToken.value_int64 = (char)CONVERT_SOURCE;
		else
			aResultToken.value_int64 = (unsigned char)CONVERT_SOURCE;
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
	if (err = UnzipItemToBuffer(huz, buf, ze.UncompressedSize, &ze))
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
	_o_throw(ERR_FAILED);
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
	if (auto aErrCode = ZipCreateBuffer(&hz, 0, (ULONGLONG)TokenToInt64(*aParam[0]), CStringCharFromTChar(ParamIndexToOptionalString(1))))
		_o_throw_zip(aErrCode);
	_o_return((__int64)hz);
}

BIF_DECL(BIF_ZipCreateFile)
{
	HZIP hz;
	if (auto aErrCode = ZipCreateFile(&hz, TokenToString(*aParam[0]), CStringCharFromTChar(ParamIndexToOptionalString(1))))
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

	if (auto obj = TokenToObject(**aParam))
	{
		GetBufferObjectPtr(aResultToken, obj, ptr, size);
		if (aResultToken.Exited())
			return;
		pwd = ParamIndexIsOmittedOrEmpty(1) ? NULL : TokenToString(*aParam[1]);
	}
	else
	{
		ptr = (size_t)TokenToInt64(*aParam[0]);
		size = (DWORD)TokenToInt64(*aParam[1]);
		pwd = ParamIndexIsOmittedOrEmpty(2) ? NULL : TokenToString(*aParam[2]);
	}
	if (ptr < 65536 || !size)
		_o_throw_value(ERR_INVALID_VALUE);
	if (size = CompressBuffer((BYTE *)ptr, buf, (DWORD)size, pwd))
		_o_return(BufferObject::Create(buf, size));
	_o_throw_oom;
}
#undef _o_throw_zip



//
// sizeof_maxsize - get max size of union or structure, helper function for BIF_Sizeof and BIF_Struct
//
BYTE sizeof_maxsize(TCHAR* buf)
{
	ResultToken ResultToken;
	int align = 0;
	ExprTokenType Var1, Var2, Var3;
	Var1.symbol = SYM_STRING;
	TCHAR tempbuf[LINE_SIZE];
	Var1.marker = tempbuf;
	Var2.symbol = SYM_INTEGER;
	Var2.value_int64 = 0;
	// used to pass aligntotal counter to structure in structure
	Var3.symbol = SYM_INTEGER;
	Var3.value_int64 = (__int64)&align;
	ExprTokenType* param[] = { &Var1, &Var2, &Var3 };
	int depth = 0;
	int max = 0, thissize = 0;
	LPCTSTR comments = 0;
	for (int i = 0; buf[i];)
	{
		if (buf[i] == '{')
		{
			depth++;
			i++;
		}
		else if (buf[i] == '}')
		{
			if (--depth < 1)
				break;
			i++;
		}
		else if (StrChrAny(&buf[i], _T(";, \t\n\r")) == &buf[i])
		{
			i++;
			continue;
		}
		else if (comments = RegExMatch(&buf[i], _T("i)^(union|struct)?\\s*(\\{|\\})\\s*\\K")))
		{
			i += (int)(comments - &buf[i]);
			continue;
		}
		else if (buf[i] == '\\' && buf[i + 1] == '\\')
		{
			if (!(comments = StrChrAny(&buf[i], _T("\r\n"))))
				break; // end of structure
			i += (int)(comments - &buf[i]);
			continue;
		}
		else
		{
			if (StrChrAny(&buf[i], _T(";,")))
			{
				_tcsncpy(tempbuf, &buf[i], thissize = (int)(StrChrAny(&buf[i], _T(";,}")) - &buf[i]));
				i += thissize + (buf[i + thissize] == '}' ? 0 : 1);
			}
			else
			{
				_tcscpy(tempbuf, &buf[i]);
				thissize = (int)_tcslen(&buf[i]);
				i += thissize;
			}
			*(tempbuf + thissize) = '\0';
			if (StrChrAny(tempbuf, _T("\t ")))
			{
				align = 0;
				BIF_sizeof(ResultToken, param, 3);
				if (max < align)
					max = (BYTE)align;
			}
			else if (max < 4)
				max = 4;
		}
	}
	return max;
}

//
// BIF_sizeof - sizeof() for structures and default types
//

BIF_DECL(BIF_sizeof)
// This code is very similar to BIF_Struct so should be maintained together
{
	int ptrsize = sizeof(UINT_PTR);		// Used for pointers on 32/64 bit system
	int offset = 0;						// also used to calculate total size of structure
	int mod = 0;
	int arraydef = 0;					// count arraysize to update offset
	int unionoffset[16] = { 0 };		// backup offset before we enter union or structure
	int unionsize[16] = { 0 };			// calculate unionsize
	bool unionisstruct[16] = { 0 };		// updated to move offset for structure in structure
	int structalign[16] = { 0 };		// keep track of struct alignment
	int totalunionsize = 0;				// total size of all unions and structures in structure
	int uniondepth = 0;					// count how deep we are in union/structure
	int align = 1;
	int* aligntotal = &align;			// pointer alignment for total structure
	int thissize = 0;					// used to save size returned from IsDefaultType
	int maxsize = 0;					// max size of union or struct
	int toalign = 0;					// custom alignment
	int thisalign = 0;

	// following are used to find variable and also get size of a structure defined in variable
	// this will hold the variable reference and offset that is given to size() to align if necessary in 64-bit
	ResultToken ResultToken;
	ExprTokenType Var1, Var2, Var3, Var4;
	Var1.symbol = SYM_VAR;
	Var2.symbol = SYM_INTEGER;

	// used to pass aligntotal counter to structure in structure
	Var3.symbol = SYM_INTEGER;
	Var3.value_int64 = (__int64)&align;
	Var4.symbol = SYM_INTEGER;
	ExprTokenType* param[] = { &Var1,&Var2,&Var3,&Var4 };

	// will hold pointer to structure definition while we parse it
	TCHAR* buf;
	size_t buf_size;
	// Should be enough buffer to accept any definition and name.
	TCHAR tempbuf[LINE_SIZE]; // just in case if we have a long comment

	// definition and field name are same max size as variables
	// also add enough room to store pointers (**) and arrays [1000]
	// give more room to use local or static variable Function(variable)
	// Parameter passed to IsDefaultType needs to be ' Definition '
	// this is because spaces are used as delimiters ( see IsDefaultType function )
	TCHAR defbuf[MAX_VAR_NAME_LENGTH * 2 + 40] = _T(" UInt "); // Set default UInt definition

	// buffer for arraysize + 2 for bracket ] and terminating character
	TCHAR intbuf[MAX_INTEGER_LENGTH + 2];

	LPTSTR bitfield = NULL;
	BYTE bitsize = 0;
	BYTE bitsizetotal = 0;
	LPTSTR isBit;

	// Set result to empty string to identify error
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");

	// if first parameter is an object (struct), simply return its size
	if (TokenToObject(*aParam[0]))
	{
		aResultToken.symbol = SYM_INTEGER;
		Struct* obj = (Struct*)TokenToObject(*aParam[0]);
		aResultToken.value_int64 = obj->mStructSize + (aParamCount > 1 ? TokenToInt64(*aParam[1]) : 0);
		return;
	}

	if (aParamCount > 1 && TokenIsNumeric(*aParam[1]))
	{	// an offset was given, set starting offset 
		offset = (int)TokenToInt64(*aParam[1]);
		Var2.value_int64 = (__int64)offset;
	}
	if (aParamCount > 2 && TokenIsNumeric(*aParam[2]))
	{   // a pointer was given to return memory to align
		aligntotal = (int*)TokenToInt64(*aParam[2]);
		Var3.value_int64 = (__int64)aligntotal;
	}
	// Set buf to beginning of structure definition
	buf = TokenToString(*aParam[0]);

	toalign = ATOI(buf);
	TCHAR alignbuf[MAX_INTEGER_LENGTH];
	if (*(buf + _tcslen(ITOA(toalign, alignbuf))) != ':' || toalign <= 0)
		Var4.value_int64 = toalign = 0;
	else
	{
		buf += _tcslen(alignbuf) + 1;
		Var4.value_int64 = toalign;
	}

	if (aParamCount > 3 && TokenIsPureNumeric(*aParam[3]))
	{   // a pointer was given to return memory to align
		toalign = (int)TokenToInt64(*aParam[3]);
		Var4.value_int64 = (__int64)toalign;
	}

	// continue as long as we did not reach end of string / structure definition
	while (*buf)
	{
		if (!_tcsncmp(buf, _T("//"), 2)) // exclude comments
		{
			buf = StrChrAny(buf, _T("\n\r")) ? StrChrAny(buf, _T("\n\r")) : (buf + _tcslen(buf));
			if (!*buf)
				break; // end of definition reached
		}
		if (buf == StrChrAny(buf, _T("\n\r\t ")))
		{	// Ignore spaces, tabs and new lines before field definition
			buf++;
			continue;
		}
		else if (_tcschr(buf, '{') && (!StrChrAny(buf, _T(";,")) || _tcschr(buf, '{') < StrChrAny(buf, _T(";,"))))
		{   // union or structure in structure definition
			if (!uniondepth++)
				totalunionsize = 0; // done here to reduce code
			if (_tcsstr(buf, _T("struct")) && _tcsstr(buf, _T("struct")) < _tcschr(buf, '{'))
				unionisstruct[uniondepth] = true; // mark that union is a structure
			else
				unionisstruct[uniondepth] = false;
			// backup offset because we need to set it back after this union / struct was parsed
			// unionsize is initialized to 0 and buffer moved to next character
			if (mod = offset % STRUCTALIGN((maxsize = sizeof_maxsize(buf))))
				offset += (thisalign - mod) % thisalign;
			structalign[uniondepth] = *aligntotal > thisalign ? STRUCTALIGN(*aligntotal) : thisalign;
			*aligntotal = 0;
			unionoffset[uniondepth] = offset; // backup offset 
			unionsize[uniondepth] = 0;
			bitsizetotal = bitsize = 0;
			// ignore even any wrong input here so it is even {mystructure...} for struct and  {anyother string...} for union
			buf = _tcschr(buf, '{') + 1;
			continue;
		}
		else if (*buf == '}')
		{	// update union
			// update size of union in case it was not updated below (e.g. last item was a union or struct)
			if ((maxsize = offset - unionoffset[uniondepth]) > unionsize[uniondepth])
				unionsize[uniondepth] = maxsize;
			// restore offset even if we had a structure in structure
			if (uniondepth > 1 && unionisstruct[uniondepth - 1])
			{
				if (mod = offset % structalign[uniondepth])
					offset += (structalign[uniondepth] - mod) % structalign[uniondepth];
			}
			else
				offset = unionoffset[uniondepth];
			if (structalign[uniondepth] > *aligntotal)
				*aligntotal = structalign[uniondepth];
			if (unionsize[uniondepth] > totalunionsize)
				totalunionsize = unionsize[uniondepth];
			// last item in union or structure, update offset now if not struct, for struct offset is up to date
			if (--uniondepth == 0)
			{
				// end of structure, align it
				//if (mod = totalunionsize % *aligntotal)
				//	totalunionsize += (*aligntotal - mod) % *aligntotal;
				// correct offset
				offset += totalunionsize;
			}
			bitsizetotal = bitsize = 0;
			buf++;
			if (buf == StrChrAny(buf, _T(";,")))
				buf++;
			continue;
		}
		// set default
		arraydef = 0;

		// copy current definition field to temporary buffer
		if (StrChrAny(buf, _T("};,")))
		{
			if ((buf_size = _tcscspn(buf, _T("};,"))) > LINE_SIZE - 1)
			{
				g_script->RuntimeError(ERR_INVALID_STRUCT, buf);
				return;
			}
			_tcsncpy(tempbuf, buf, buf_size);
			tempbuf[buf_size] = '\0';
		}
		else if (_tcslen(buf) > LINE_SIZE - 1)
		{
			g_script->RuntimeError(ERR_INVALID_STRUCT, buf);
			return;
		}
		else
			_tcscpy(tempbuf, buf);

		// Trim trailing spaces
		rtrim(tempbuf);

		// Array
		if (_tcschr(tempbuf, '['))
		{
			_tcsncpy(intbuf, _tcschr(tempbuf, '['), MAX_INTEGER_LENGTH);
			intbuf[_tcscspn(intbuf, _T("]")) + 1] = '\0';
			arraydef = (int)ATOI64(intbuf + 1);
			// remove array definition
			StrReplace(tempbuf, intbuf, _T(""), SCS_SENSITIVE, 1, LINE_SIZE);
		}
		if (_tcschr(tempbuf, '['))
		{	// array to array and similar not supported
			g_script->RuntimeError(ERR_INVALID_STRUCT, tempbuf);
			return;
		}
		// Pointer, while loop will continue here because we only need size
		if (_tcschr(tempbuf, '*'))
		{
			if (_tcschr(tempbuf, ':'))
			{
				g_script->RuntimeError(ERR_INVALID_STRUCT_BIT_POINTER, tempbuf);
				return;
			}
			// align offset for pointer
			if (mod = offset % STRUCTALIGN(ptrsize))
				offset += (thisalign - mod) % thisalign;
			offset += ptrsize * (arraydef ? arraydef : 1);
			if (thisalign > *aligntotal)
				*aligntotal = thisalign;
			// update offset
			if (uniondepth)
			{
				if ((maxsize = offset - unionoffset[uniondepth]) > unionsize[uniondepth])
					unionsize[uniondepth] = maxsize;
				// reset offset if in union and union is not a structure
				if (!unionisstruct[uniondepth])
					offset = unionoffset[uniondepth];
			}

			// Move buffer pointer now and continue
			if (_tcschr(buf, '}') && (!StrChrAny(buf, _T(";,")) || _tcschr(buf, '}') < StrChrAny(buf, _T(";,"))))
				buf += _tcscspn(buf, _T("}")); // keep } character to update union
			else if (StrChrAny(buf, _T(";,")))
				buf += _tcscspn(buf, _T(";,")) + 1;
			else
				buf += _tcslen(buf);
			continue;
		}

		// if offset is 0 and there are no };, characters, it means we have a pure definition
		if (StrChrAny(tempbuf, _T(" \t")) || StrChrAny(tempbuf, _T("};,")) || (!StrChrAny(buf, _T("};,")) && !offset))
		{
			if ((buf_size = _tcscspn(tempbuf, _T("\t ["))) > MAX_VAR_NAME_LENGTH * 2 + 30)
			{
				g_script->RuntimeError(ERR_INVALID_STRUCT, tempbuf);
				return;
			}
			isBit = StrChrAny(omit_leading_whitespace(tempbuf), _T(" \t"));
			if (!isBit || *isBit != ':')
			{
				if (_tcsnicmp(defbuf + 1, tempbuf, _tcslen(defbuf) - 2))
					bitsizetotal = bitsize = 0;
				_tcsncpy(defbuf + 1, tempbuf, _tcscspn(tempbuf, _T("\t [")));
				_tcscpy(defbuf + 1 + _tcscspn(tempbuf, _T("\t [")), _T(" "));
			}
			if (bitfield = _tcschr(tempbuf, ':'))
			{
				if (bitsizetotal / 8 == thissize)
					bitsizetotal = bitsize = 0;
				bitsizetotal += bitsize = ATOI(bitfield + 1);
			}
			else
				bitsizetotal = bitsize = 0;
		}
		else // Not 'TypeOnly' definition because there are more than one fields in structure so use default type UInt
		{
			// Commented out following line to keep previous or default UInt definition like in c++, e.g. "Int x,y,Char a,b", 
			// Note: separator , or ; can be still used but
			// _tcscpy(defbuf,_T(" UInt "));
			if (bitfield = _tcschr(tempbuf, ':'))
			{
				if (bitsizetotal / 8 == thissize)
					bitsizetotal = bitsize = 0;
				bitsizetotal += bitsize = ATOI(bitfield + 1);
			}
			else
				bitsizetotal = bitsize = 0;
		}

		// Now find size in default types array and create new field
		// If Type not found, resolve type to variable and get size of struct defined in it
		if ((!_tcscmp(defbuf, _T(" bool ")) && (thissize = 1)) || (thissize = IsDefaultType(defbuf)))
		{
			// align offset
			if (!bitsize || bitsizetotal == bitsize)
			{
				if (thissize == 1)
					thisalign = 1;
				else if (mod = offset % STRUCTALIGN(thissize))
					offset += (thisalign - mod) % thisalign;
				offset += thissize * (arraydef ? arraydef : 1);
			}
			if (thisalign > *aligntotal)
				*aligntotal = thisalign; // > ptrsize ? ptrsize : thissize;
		}
		else // type was not found, check for user defined type in variables
		{
			Var1.var = NULL;
			UserFunc* bkpfunc = NULL;
			// check if we have a local/static declaration and resolve to function
			// For example Struct("MyFunc(mystruct) mystr")
			if (_tcschr(defbuf, '('))
			{
				bkpfunc = g->CurrentFunc; // don't bother checking, just backup and restore later
				g->CurrentFunc = (UserFunc*)g_script->FindGlobalFunc(defbuf + 1, _tcscspn(defbuf, _T("(")) - 1);
				if (g->CurrentFunc) // break if not found to identify error
				{
					_tcscpy(tempbuf, defbuf + 1);
					_tcscpy(defbuf + 1, tempbuf + _tcscspn(tempbuf, _T("(")) + 1); //,_tcschr(tempbuf,')') - _tcschr(tempbuf,'('));
					_tcscpy(_tcschr(defbuf, ')'), _T(" \0"));
					Var1.var = g_script->FindVar(defbuf + 1, _tcslen(defbuf) - 2, FINDVAR_LOCAL);
					g->CurrentFunc = bkpfunc;
				}
				else // release object and return
				{
					g->CurrentFunc = bkpfunc;
					g_script->RuntimeError(ERR_INVALID_STRUCT_IN_FUNC, defbuf);
					return;
				}
			}
			else if (g->CurrentFunc) // try to find local variable first
				Var1.var = g_script->FindVar(defbuf + 1, _tcslen(defbuf) - 2, FINDVAR_LOCAL);
			// try to find global variable if local was not found or we are not in func
			if (Var1.var == NULL)
				Var1.var = g_script->FindVar(defbuf + 1, _tcslen(defbuf) - 2, FINDVAR_GLOBAL);
			if (Var1.var != NULL)
			{
				// Call BIF_sizeof passing offset in second parameter to align if necessary
				int newaligntotal = sizeof_maxsize(TokenToString(Var1));
				if (STRUCTALIGN(newaligntotal) > *aligntotal)
					*aligntotal = thisalign;
				if ((!bitsize || bitsizetotal == bitsize) && offset && (mod = offset % *aligntotal))
					offset += (*aligntotal - mod) % *aligntotal;
				param[1]->value_int64 = (__int64)offset;
				BIF_sizeof(ResultToken, param, 3);
				if (ResultToken.symbol != SYM_INTEGER)
				{	// could not resolve structure
					g_script->RuntimeError(ERR_INVALID_STRUCT, defbuf);
					return;
				}
				// sizeof was given an offset that it applied and aligned if necessary, so set offset =  and not +=
				if (!bitsize || bitsizetotal == bitsize)
					offset = (int)ResultToken.value_int64 + (arraydef ? ((arraydef - 1) * ((int)ResultToken.value_int64 - offset)) : 0);
			}
			else // No variable was found and it is not default type so we can't determine size, return empty string.
			{
				g_script->RuntimeError(ERR_INVALID_STRUCT, defbuf);
				return;
			}
		}
		// update union size
		if (uniondepth)
		{
			if ((maxsize = offset - unionoffset[uniondepth]) > unionsize[uniondepth])
				unionsize[uniondepth] = maxsize;
			if (!unionisstruct[uniondepth])
			{
				// reset offset if in union and union is not a structure
				offset = unionoffset[uniondepth];
				// reset bit offset and size
				bitsize = bitsizetotal = 0;
			}
		}
		// Move buffer pointer now
		if (_tcschr(buf, '}') && (!StrChrAny(buf, _T(";,")) || _tcschr(buf, '}') < StrChrAny(buf, _T(";,"))))
			buf += _tcscspn(buf, _T("}")); // keep } character to update union
		else if (StrChrAny(buf, _T(";,")))
			buf += _tcscspn(buf, _T(";,")) + 1;
		else
			buf += _tcslen(buf);
	}
	if (*aligntotal && (mod = offset % *aligntotal)) // align even if offset was given e.g. for _NMHDR:="HWND hwndFrom,UINT_PTR idFrom,UINT code", _NMTVGETINFOTIP: = "_NMHDR hdr,UINT uFlags,UInt link"
		offset += (*aligntotal - mod) % *aligntotal;
	aResultToken.symbol = SYM_INTEGER;
	aResultToken.value_int64 = offset;
}

//
// BIF_Struct - Create structure
//

BIF_DECL(BIF_Struct)
{
	// At least the definition for structure must be given
	if (!aParamCount)
		return;
	Struct* obj = Struct::Create(aParam, aParamCount);
	if (obj)
	{
		aResultToken.symbol = SYM_OBJECT;
		aResultToken.object = obj;
		return;
		// DO NOT ADDREF: after we return, the only reference will be in aResultToken.
	}
	else
	{
		aResultToken.SetResult(FAIL);
		return;
	}
}

//
// ObjRawSize()
//

__int64 ObjRawSize(IObject* aObject, IObject* aObjects)
{
	ResultToken result_token, this_token, aKey, aValue;
	ExprTokenType* params[] = { &aKey, &aValue };
	this_token.symbol = SYM_OBJECT;

	if ((_tcscmp(aObject->Type(), _T("Object")))
		&& (_tcscmp(aObject->Type(), _T("Array")))
		&& (_tcscmp(aObject->Type(), _T("Map")))
		&& (_tcscmp(aObject->Type(), _T("Buffer")))
		&& (_tcscmp(aObject->Type(), _T("Struct")))
		&& (_tcscmp(aObject->Type(), _T("ComValue"))))
		g_script->RuntimeError(ERR_TYPE_MISMATCH, aObject->Type());

	__int64 aSize = 1 + sizeof(__int64);
	result_token.InitResult(L"");
	if (!_tcscmp(aObject->Type(), _T("Buffer")))
	{
		this_token.object = aObject;
		result_token.InitResult(L"");
		aObject->Invoke(result_token, IT_GET, _T("size"), this_token, nullptr, 0);
		aSize += result_token.value_int64;
	}
	else if (!_tcscmp(aObject->Type(), _T("ComValue")))
	{
		ComObject* obj = dynamic_cast<ComObject*>(aObject);
		if (!obj || !(obj->mVarType < 9 || obj->mVarType - 16 < 8))
			g_script->RuntimeError(ERR_TYPE_MISMATCH, _T("Only following values are supported VT_EMPTY/NULL/BOOL/I2/I4/R4/R8/CY/DATE/BSTR/I1/UI1/UI2/UI4/I8/UI8/INT/UINT"));
		aSize += 1 + (obj->mVarType == VT_BSTR ? SysStringByteLen((BSTR)obj->mValPtr) + sizeof(TCHAR) : sizeof(__int64));
	}

	this_token.object = aObjects;

	IObject* enumerator;
	ResultType result;
	ResultToken enum_token;
	enum_token.object = aObject;
	enum_token.symbol = SYM_OBJECT;

	// create variables to use in for loop / for enumeration
	// these will be deleted afterwards

	Var vkey, vval;

	aKey.symbol = SYM_OBJECT;
	aKey.object = aObject;
	this_token.object = aObjects;
	result_token.InitResult(L"");
	aObjects->Invoke(result_token, IT_CALL, _T("Has"), this_token, params, 1);
	if (!result_token.value_int64)
	{
		aKey.symbol = SYM_OBJECT;
		aKey.object = aObject;
		aValue.symbol = SYM_STRING;
		aValue.marker = _T("");
		aValue.marker_length = 0;
		result_token.InitResult(L"");
		aObjects->Invoke(result_token, IT_SET, 0, this_token, params, 2);
	}

	if (!_tcscmp(aObject->Type(), _T("ComValue")))
		return aSize;

	// Prepare parameters for the loop below: enum.Next(var1 [, var2])
	aKey.SetVarRef(&vkey);
	aValue.SetVarRef(&vval);

	IObject* aIsObject;
	__int64 aIsValue;
	SymbolType aVarType;
	if (_tcscmp(aObject->Type(), _T("Buffer")) && _tcscmp(aObject->Type(), _T("Object")) && _tcscmp(aObject->Type(), _T("ComValue")))
	{
		result = GetEnumerator(enumerator, enum_token, 2, false);
		// Check if object returned an enumerator, otherwise return
		if (result != OK)
			return 0;
		for (;;)
		{
			result = CallEnumerator(enumerator, params, 2, false);
			if (result != CONDITION_TRUE)
				break;

			if (aIsObject = TokenToObject(aKey))
			{
				this_token.object = aObjects;
				result_token.InitResult(L"");
				aObjects->Invoke(result_token, IT_CALL, _T("Has"), this_token, params, 1);
				if (result_token.value_int64)
					aSize += 1 + sizeof(__int64);
				else
					aSize += ObjRawSize(aIsObject, aObjects);
			}
			else
			{
				vkey.ToTokenSkipAddRef(aKey);
				if ((aVarType = aKey.symbol) == SYM_STRING)
					aSize += (aKey.marker_length ? (aKey.marker_length + 1) * sizeof(TCHAR) : 0) + 9;
				else
					aSize += (aVarType == SYM_FLOAT || (aIsValue = TokenToInt64(aKey)) > 4294967295) ? 9 : aIsValue > 65535 ? 5 : aIsValue > 255 ? 3 : aIsValue > -129 ? 2 : aIsValue > -32769 ? 3 : aIsValue >= INT_MIN ? 5 : 9;
			}

			if (aIsObject = TokenToObject(aValue))
			{
				aKey.symbol = SYM_OBJECT;
				aKey.object = aIsObject;
				this_token.object = aObjects;
				result_token.InitResult(L"");
				aObjects->Invoke(result_token, IT_CALL, _T("Has"), this_token, params, 1);
				if (result_token.value_int64)
					aSize += 1 + sizeof(__int64);
				else
					aSize += ObjRawSize(aIsObject, aObjects);
			}
			else
			{
				aValue.var->ToTokenSkipAddRef(aValue);
				if ((aVarType = aValue.symbol) == SYM_STRING)
					aSize += (aValue.marker_length ? (aValue.marker_length + 1) * sizeof(TCHAR) : 0) + 9;
				else
					aSize += aVarType == SYM_FLOAT || (aIsValue = TokenToInt64(aValue)) > 4294967295 ? 9 : aIsValue > 65535 ? 5 : aIsValue > 255 ? 3 : aIsValue > -129 ? 2 : aIsValue > -32769 ? 3 : aIsValue >= INT_MIN ? 5 : 9;
			}

			// release object if it was assigned prevoiously when calling enum.Next
			if (vkey.IsObject())
				vkey.ReleaseObject();
			if (vval.IsObject())
				vval.ReleaseObject();

			aKey.SetVarRef(&vkey);
			aValue.SetVarRef(&vval);
		}
		enumerator->Release();
	}

	result_token.InitResult(L"");
	this_token.object = aObject;
	aObject->Invoke(result_token, IT_CALL, _T("OwnProps"), this_token, nullptr, 0);
	if (result_token.symbol == SYM_OBJECT)
	{
		enumerator = result_token.object;
		for (;;)
		{
			result = CallEnumerator(enumerator, params, 2, false);
			if (result != CONDITION_TRUE)
				break;

			// Properties are always strings
			vkey.ToTokenSkipAddRef(aKey);
			aSize += (aKey.marker_length ? (aKey.marker_length + 1) * sizeof(TCHAR) : 0) + 9;

			if (aIsObject = TokenToObject(aValue))
			{
				aKey.symbol = SYM_OBJECT;
				aKey.object = aIsObject;
				this_token.object = aObjects;
				result_token.InitResult(L"");
				aObjects->Invoke(result_token, IT_CALL, _T("Has"), this_token, params, 1);
				if (result_token.value_int64)
					aSize += 1 + sizeof(__int64);
				else
					aSize += ObjRawSize(aIsObject, aObjects);
			}
			else
			{
				aValue.var->ToTokenSkipAddRef(aValue);
				if ((aVarType = aValue.symbol) == SYM_STRING)
					aSize += (aValue.marker_length ? (aValue.marker_length + 1) * sizeof(TCHAR) : 0) + 9;
				else
					aSize += aVarType == SYM_FLOAT || (aIsValue = TokenToInt64(aValue)) > 4294967295 ? 9 : aIsValue > 65535 ? 5 : aIsValue > 255 ? 3 : aIsValue > -129 ? 2 : aIsValue > -32769 ? 3 : aIsValue >= INT_MIN ? 5 : 9;
			}

			// release object if it was assigned prevoiously when calling enum.Next
			if (vkey.IsObject())
				vkey.ReleaseObject();
			if (vval.IsObject())
				vval.ReleaseObject();

			aKey.symbol = SYM_VAR;
			aKey.var = &vkey;
			aKey.var_usage = VARREF_REF;
			aValue.symbol = SYM_VAR;
			aValue.var = &vval;
			aValue.var_usage = VARREF_REF;
		}
		enumerator->Release();
	}
	vkey.Free();
	vval.Free();
	return aSize;
}

//
// ObjRawDump()
//

__int64 ObjRawDump(IObject* aObject, char* aBuffer, Map* aObjects, UINT& aObjCount)
{

	ResultToken result_token, this_token, aKey, aValue;
	ExprTokenType* params[] = { &aKey, &aValue };

	this_token.symbol = SYM_OBJECT;

	Object* aObj = dynamic_cast<Object*>(aObject);
	char* aThisBuffer = aBuffer;
	size_t bufsize;
	if (!_tcscmp(aObject->Type(), _T("Object")))
		*aThisBuffer = (char)(aObj->IsUnsorted() ? -14 : -13);
	else if (!_tcscmp(aObject->Type(), _T("Array")))
		*aThisBuffer = (char)(aObj->IsUnsorted() ? -16 : -15);
	else if (!_tcscmp(aObject->Type(), _T("Map")))
		*aThisBuffer = (char)(aObj->IsUnsorted() ? -18 : -17);
	else if (!_tcscmp(aObject->Type(), _T("Struct")))
		*aThisBuffer = (char)-18;
	else if (!_tcscmp(aObject->Type(), _T("Buffer")))
	{
		this_token.object = aObject;
		result_token.InitResult(L"");
		aObject->Invoke(result_token, IT_GET, _T("size"), this_token, nullptr, 0);
		bufsize = (size_t)result_token.value_int64;
		*aThisBuffer = (char)-19;
		result_token.InitResult(L"");
		aObject->Invoke(result_token, IT_GET, _T("ptr"), this_token, nullptr, 0);
		memmove(aThisBuffer + 1 + sizeof(__int64), (void*)result_token.value_int64, (size_t)bufsize);
		*(__int64*)(aThisBuffer + 1) = bufsize;
		aThisBuffer += bufsize;
	}
	else if (!_tcscmp(aObject->Type(), _T("ComValue")))
	{
		ComObject* obj = dynamic_cast<ComObject*>(aObject);
		if (!obj || !(obj->mVarType < 9 || obj->mVarType - 16 < 8))
			g_script->RuntimeError(ERR_TYPE_MISMATCH, _T("Only following values are supported VT_EMPTY/NULL/BOOL/I2/I4/R4/R8/CY/DATE/BSTR/I1/UI1/UI2/UI4/I8/UI8/INT/UINT"));
		*aThisBuffer = (char)-20;
		*(aThisBuffer + 1 + sizeof(__int64)) = (char)obj->mVarType;
		if (obj->mVarType == VT_BSTR)
		{
			bufsize = SysStringByteLen((BSTR)obj->mValPtr) + sizeof(TCHAR);
			*(__int64*)(aThisBuffer + 1) = 1 + bufsize;
			memmove(aThisBuffer + 2 + sizeof(__int64), obj->mValPtr, bufsize);
			aThisBuffer += 1 + bufsize;
		}
		else
		{
			*(__int64*)(aThisBuffer + 1) = 1 + sizeof(__int64);
			*(__int64*)(aThisBuffer + 2 + sizeof(__int64)) = obj->mVal64;
			aThisBuffer += 1 + sizeof(__int64);
		}
	}
	else
		g_script->RuntimeError(ERR_TYPE_MISMATCH, aObject->Type());

	aThisBuffer += 1 + sizeof(__int64);

	IObject* enumerator;
	ResultType result;
	Var vkey, vval;

	aKey.symbol = SYM_OBJECT;
	aKey.object = aObject;
	this_token.object = aObjects;
	result_token.InitResult(L"");
	aObjects->Invoke(result_token, IT_CALL, _T("Has"), this_token, params, 1);
	if (!result_token.value_int64)
	{
		aValue.symbol = SYM_INTEGER;
		aValue.value_int64 = aObjCount++;
		result_token.InitResult(L"");
		aObjects->Invoke(result_token, IT_SET, 0, this_token, params, 2);
	}

	if (!_tcscmp(aObject->Type(), _T("ComValue")))
		return aThisBuffer - aBuffer;

	// Prepare parameters for the loop below
	aKey.symbol = SYM_VAR;
	aKey.var = &vkey;
	aKey.mem_to_free = 0;
	aKey.var_usage = VARREF_REF;
	aValue.symbol = SYM_VAR;
	aValue.var = &vval;
	aValue.mem_to_free = 0;
	aValue.var_usage = VARREF_REF;

	IObject* aIsObject;
	__int64 aIsValue;
	__int64 aThisSize;
	SymbolType aVarType;

	if (_tcscmp(aObject->Type(), _T("Buffer")) && _tcscmp(aObject->Type(), _T("Object")))
	{
		result_token.object = aObject;
		result_token.symbol = SYM_OBJECT;
		result = GetEnumerator(enumerator, result_token, 2, false);
		if (result != OK)
			return NULL;
		for (;;)
		{
			result = CallEnumerator(enumerator, params, 2, false);
			if (result != CONDITION_TRUE)
				break;

			// copy Key
			if (aIsObject = TokenToObject(aKey))
			{
				result_token.InitResult(L"");
				aObjects->Invoke(result_token, IT_CALL, _T("Has"), this_token, params, 1);
				if (result_token.value_int64)
				{
					result_token.InitResult(L"");
					aObjects->Invoke(result_token, IT_GET, 0, this_token, params, 1);
					*aThisBuffer = (char)-12;
					aThisBuffer += 1;
					*(__int64*)aThisBuffer = result_token.value_int64;
					aThisBuffer += sizeof(__int64);
				}
				else
				{
					aThisSize = ObjRawDump(aIsObject, aThisBuffer, aObjects, aObjCount);
					*(__int64*)(aThisBuffer + 1) = aThisSize - 1 - sizeof(__int64);
					aThisBuffer += aThisSize;
				}
			}
			else
			{
				aKey.var->ToTokenSkipAddRef(aKey);
				if ((aVarType = aKey.symbol) == SYM_STRING)
				{
					*aThisBuffer++ = (char)-10;
					*(__int64*)aThisBuffer = aThisSize = (__int64)(aKey.marker_length ? (aKey.marker_length + 1) * sizeof(TCHAR) : 0);
					aThisBuffer += sizeof(__int64);
					if (aThisSize)
					{
						memcpy(aThisBuffer, aKey.marker, (size_t)aThisSize);
						aThisBuffer += aThisSize;
					}
				}
				else if (aVarType == SYM_FLOAT)
				{
					*aThisBuffer++ = (char)-9;
					*(double*)aThisBuffer = TokenToDouble(aKey);
					aThisBuffer += sizeof(__int64);
				}
				else if ((aIsValue = TokenToInt64(aKey)) > 4294967295)
				{
					*aThisBuffer++ = (char)-8;
					*(__int64*)aThisBuffer = aIsValue;
					aThisBuffer += sizeof(__int64);
				}
				else if (aIsValue > 65535)
				{
					*aThisBuffer++ = (char)-6;
					*(UINT*)aThisBuffer = (UINT)aIsValue;
					aThisBuffer += sizeof(UINT);
				}
				else if (aIsValue > 255)
				{
					*aThisBuffer++ = (char)-4;
					*(USHORT*)aThisBuffer = (USHORT)aIsValue;
					aThisBuffer += sizeof(USHORT);
				}
				else if (aIsValue > -1)
				{
					*aThisBuffer++ = (char)-2;
					*aThisBuffer = (BYTE)aIsValue;
					aThisBuffer += sizeof(BYTE);
				}
				else if (aIsValue > -129)
				{
					*aThisBuffer++ = (char)-1;
					*aThisBuffer = (char)aIsValue;
					aThisBuffer += sizeof(char);
				}
				else if (aIsValue > -32769)
				{
					*aThisBuffer++ = (char)-3;
					*(short*)aThisBuffer = (short)aIsValue;
					aThisBuffer += sizeof(short);
				}
				else if (aIsValue >= INT_MIN)
				{
					*aThisBuffer++ = (char)-5;
					*(int*)aThisBuffer = (int)aIsValue;
					aThisBuffer += sizeof(int);
				}
				else
				{
					*aThisBuffer++ = (char)-7;
					*(__int64*)aThisBuffer = (__int64)aIsValue;
					aThisBuffer += sizeof(__int64);
				}
			}

			// copy Value
			if (aIsObject = TokenToObject(aValue))
			{
				aKey.symbol = SYM_OBJECT;
				aKey.object = aIsObject;
				result_token.InitResult(L"");
				aObjects->Invoke(result_token, IT_CALL, _T("Has"), this_token, params + 1, 1);
				if (result_token.value_int64)
				{
					result_token.InitResult(L"");
					aObjects->Invoke(result_token, IT_GET, 0, this_token, params + 1, 1);
					*aThisBuffer++ = (char)12;
					*(__int64*)aThisBuffer = result_token.value_int64;
					aThisBuffer += sizeof(__int64);
				}
				else
				{
					aThisSize = ObjRawDump(aIsObject, aThisBuffer, aObjects, aObjCount);
					*(__int64*)(aThisBuffer + 1) = aThisSize - 1 - sizeof(__int64);
					*(char*)aThisBuffer = -1 * *(char*)aThisBuffer;
					aThisBuffer += aThisSize;
				}
			}
			else
			{
				aValue.var->ToTokenSkipAddRef(aValue);
				if ((aVarType = aValue.symbol) == SYM_STRING)
				{
					*aThisBuffer++ = (char)10;
					*(__int64*)aThisBuffer = aThisSize = (__int64)(aValue.marker_length ? (aValue.marker_length + 1) * sizeof(TCHAR) : 0);
					aThisBuffer += sizeof(__int64);
					if (aThisSize)
					{
						memcpy(aThisBuffer, aValue.marker, (size_t)aThisSize);
						aThisBuffer += aThisSize;
					}
				}
				else if (aVarType == SYM_FLOAT)
				{
					*aThisBuffer++ = (char)9;
					*(double*)aThisBuffer = TokenToDouble(aValue);
					aThisBuffer += sizeof(double);
				}
				else if ((aIsValue = TokenToInt64(aValue)) > 4294967295)
				{
					*aThisBuffer++ = (char)8;
					*(__int64*)aThisBuffer = aIsValue;
					aThisBuffer += sizeof(__int64);
				}
				else if (aIsValue > 65535)
				{
					*aThisBuffer++ = (char)6;
					*(UINT*)aThisBuffer = (UINT)aIsValue;
					aThisBuffer += sizeof(UINT);
				}
				else if (aIsValue > 255)
				{
					*aThisBuffer++ = (char)4;
					*(USHORT*)aThisBuffer = (USHORT)aIsValue;
					aThisBuffer += sizeof(USHORT);
				}
				else if (aIsValue > -1)
				{
					*aThisBuffer++ = (char)2;
					*aThisBuffer = (BYTE)aIsValue;
					aThisBuffer += sizeof(BYTE);
				}
				else if (aIsValue > -129)
				{
					*aThisBuffer++ = (char)1;
					*aThisBuffer = (char)aIsValue;
					aThisBuffer += sizeof(char);
				}
				else if (aIsValue > -32769)
				{
					*aThisBuffer++ = (char)3;
					*(short*)aThisBuffer = (short)aIsValue;
					aThisBuffer += sizeof(short);
				}
				else if (aIsValue > INT_MIN)
				{
					*aThisBuffer++ = (char)5;
					*aThisBuffer = (int)aIsValue;
					aThisBuffer += sizeof(int);
				}
				else
				{
					*aThisBuffer++ = (char)7;
					*(__int64*)aThisBuffer = (__int64)aIsValue;
					aThisBuffer += sizeof(__int64);
				}
			}

			// release object if it was assigned prevoiously when calling enum
			if (vkey.IsObject())
				vkey.ReleaseObject();
			if (vval.IsObject())
				vval.ReleaseObject();

			aKey.symbol = SYM_VAR;
			aKey.var = &vkey;
			aKey.var_usage = VARREF_REF;
			aValue.symbol = SYM_VAR;
			aValue.var = &vval;
			aValue.var_usage = VARREF_REF;
		}
		// release enumerator and free vars
		enumerator->Release();
	}

	this_token.object = aObject;
	result_token.InitResult(L"");
	aObject->Invoke(result_token, IT_CALL, _T("OwnProps"), this_token, nullptr, 0);
	if (result_token.symbol == SYM_OBJECT)
		enumerator = result_token.object;
	else
		return NULL;
	this_token.object = aObjects;

	for (;;)
	{
		result = CallEnumerator(enumerator, params, 2, false);
		if (result != CONDITION_TRUE)
			break;

		// copy Key
		aKey.var->ToTokenSkipAddRef(aKey);
		*aThisBuffer++ = (char)-11;
		*(__int64*)aThisBuffer = aThisSize = (__int64)(aKey.marker_length ? (aKey.marker_length + 1) * sizeof(TCHAR) : 0);
		aThisBuffer += sizeof(__int64);
		if (aThisSize)
		{
			memcpy(aThisBuffer, aKey.marker, (size_t)aThisSize);
			aThisBuffer += aThisSize;
		}

		// copy Value
		if (aIsObject = TokenToObject(aValue))
		{
			aKey.symbol = SYM_OBJECT;
			aKey.object = aIsObject;
			result_token.InitResult(L"");
			aObjects->Invoke(result_token, IT_CALL, _T("Has"), this_token, params + 1, 1);
			if (result_token.value_int64)
			{
				result_token.InitResult(L"");
				aObjects->Invoke(result_token, IT_GET, 0, this_token, params + 1, 1);
				*aThisBuffer++ = (char)12;
				*(__int64*)aThisBuffer = result_token.value_int64;
				aThisBuffer += sizeof(__int64);
			}
			else
			{
				aThisSize = ObjRawDump(aIsObject, aThisBuffer, aObjects, aObjCount);
				*(__int64*)(aThisBuffer + 1) = aThisSize - 1 - sizeof(__int64);
				*(char*)aThisBuffer = -1 * *(char*)aThisBuffer;
				aThisBuffer += aThisSize;
			}
		}
		else
		{
			aValue.var->ToTokenSkipAddRef(aValue);
			if ((aVarType = aValue.symbol) == SYM_STRING)
			{
				*aThisBuffer++ = (char)10;
				*(__int64*)aThisBuffer = aThisSize = (__int64)(aValue.marker_length ? (aValue.marker_length + 1) * sizeof(TCHAR) : 0);
				aThisBuffer += sizeof(__int64);
				if (aThisSize)
				{
					memcpy(aThisBuffer, aValue.marker, (size_t)aThisSize);
					aThisBuffer += aThisSize;
				}
			}
			else if (aVarType == SYM_FLOAT)
			{
				*aThisBuffer++ = (char)9;
				*(double*)aThisBuffer = TokenToDouble(aValue);
				aThisBuffer += sizeof(double);
			}
			else if ((aIsValue = TokenToInt64(aValue)) > 4294967295)
			{
				*aThisBuffer++ = (char)8;
				*(__int64*)aThisBuffer = aIsValue;
				aThisBuffer += sizeof(__int64);
			}
			else if (aIsValue > 65535)
			{
				*aThisBuffer++ = (char)6;
				*(UINT*)aThisBuffer = (UINT)aIsValue;
				aThisBuffer += sizeof(UINT);
			}
			else if (aIsValue > 255)
			{
				*aThisBuffer++ = (char)4;
				*(USHORT*)aThisBuffer = (USHORT)aIsValue;
				aThisBuffer += sizeof(USHORT);
			}
			else if (aIsValue > -1)
			{
				*aThisBuffer++ = (char)2;
				*aThisBuffer = (BYTE)aIsValue;
				aThisBuffer += sizeof(BYTE);
			}
			else if (aIsValue > -129)
			{
				*aThisBuffer++ = (char)1;
				*aThisBuffer = (char)aIsValue;
				aThisBuffer += sizeof(char);
			}
			else if (aIsValue > -32769)
			{
				*aThisBuffer++ = (char)3;
				*(short*)aThisBuffer = (short)aIsValue;
				aThisBuffer += sizeof(short);
			}
			else if (aIsValue > INT_MIN)
			{
				*aThisBuffer++ = (char)5;
				*aThisBuffer = (int)aIsValue;
				aThisBuffer += sizeof(int);
			}
			else
			{
				*aThisBuffer++ = (char)7;
				*(__int64*)aThisBuffer = (__int64)aIsValue;
				aThisBuffer += sizeof(__int64);
			}
		}

		// release object if it was assigned prevoiously when calling enum
		if (vkey.IsObject())
			vkey.ReleaseObject();
		if (vval.IsObject())
			vval.ReleaseObject();

		aKey.symbol = SYM_VAR;
		aKey.var = &vkey;
		aKey.var_usage = VARREF_REF;
		aValue.symbol = SYM_VAR;
		aValue.var = &vval;
		aValue.var_usage = VARREF_REF;
	}
	// release enumerator and free vars
	enumerator->Release();

	vkey.Free();
	vval.Free();
	return aThisBuffer - aBuffer;
}

//
// ObjDump()
//

BIF_DECL(BIF_ObjDump)
{
	aResultToken.symbol = SYM_INTEGER;
	IObject* aObject;
	if (!(aObject = TokenToObject(*aParam[0])))
		_o_throw_value(ERR_INVALID_ARG_TYPE);
	Map* aObjects = Map::Create();
	DWORD aSize = (DWORD)ObjRawSize(aObject, aObjects);
	aObjects->Release();
	aObjects = Map::Create();
	char* aBuffer = (char*)malloc(aSize);
	if (!aBuffer)
		_o_throw_oom;
	*(__int64*)(aBuffer + 1) = aSize - 1 - sizeof(__int64);
	UINT aObjCount = 0;
	if (aSize != ObjRawDump(aObject, aBuffer, aObjects, aObjCount))
	{
		aObjects->Release();
		free(aBuffer);
		_o_throw(_T("Error dumping Object."));
	}
	aObjects->Release();
	if (aParamCount > 1)
	{
		LPVOID aDataBuf;
		DWORD aCompressedSize = CompressBuffer((BYTE*)aBuffer, aDataBuf, aSize, ParamIndexIsOmittedOrEmpty(1) ? NULL : TokenToString(*aParam[1]));
		free(aBuffer);
		if (aCompressedSize)
		{
			aObject = BufferObject::Create(aDataBuf, aCompressedSize);
			((BufferObject*)aObject)->SetBase(BufferObject::sPrototype);
		}
		else
			_o_throw_oom;
	}
	else
	{
		aObject = BufferObject::Create(aBuffer, aSize);
		dynamic_cast<BufferObject*>(aObject)->SetBase(BufferObject::sPrototype);
	}
	aResultToken.SetValue(aObject);
}

//
// ObjRawLoad()
//

IObject* ObjRawLoad(char* aBuffer, IObject**& aObjects, UINT& aObjCount, UINT& aObjSize)
{
	IObject* aObject = NULL;
	ResultToken result_token, this_token, enum_token, aKey, aValue;
	ExprTokenType* params[] = { &aKey, &aValue };
	TCHAR buf[MAX_INTEGER_LENGTH];
	result_token.buf = buf;
	aKey.mem_to_free = NULL;
	aValue.mem_to_free = NULL;
	this_token.symbol = SYM_OBJECT;

	if (aObjCount == aObjSize)
	{
		IObject** newObjects = (IObject**)malloc(aObjSize * 2 * sizeof(IObject**));
		if (!newObjects)
			return 0;
		memcpy(newObjects, aObjects, aObjSize * sizeof(IObject**));
		free(aObjects);
		aObjects = newObjects;
		aObjSize *= 2;
	}

	char* aThisBuffer = aBuffer;
	char typekey, typeval, typeobj = *(char*)aThisBuffer;
	size_t mSize;
	if (typeobj == -13 || typeobj == 13) //(_tcscmp(aObject->Type(), _T("Object")))
		aObject = Object::Create();
	else if (typeobj == -14 || typeobj == 14) //(_tcscmp(aObject->Type(), _T("Object")))
		aObject = Object::Create(nullptr, 0, nullptr, true);
	else if (typeobj == -15 || typeobj == 15) //(_tcscmp(aObject->Type(), _T("Array")))
		aObject = Array::Create();
	else if (typeobj == -16 || typeobj == 16) //(_tcscmp(aObject->Type(), _T("Array")))
		aObject = Array::Create(nullptr, 0, true);
	else if (typeobj == -17 || typeobj == 17) //(_tcscmp(aObject->Type(), _T("Map")) || _tcscmp(aObject->Type(), _T("Struct")))
		aObject = Map::Create();
	else if (typeobj == -18 || typeobj == 18) //(_tcscmp(aObject->Type(), _T("Map")) || _tcscmp(aObject->Type(), _T("Struct")))
		aObject = Map::Create(nullptr, 0, true);
	else if (typeobj == -19 || typeobj == 19) //(_tcscmp(aObject->Type(), _T("Buffer")))
	{
		mSize = (size_t) * (__int64*)(aThisBuffer + 1);
		aKey.buf = (LPTSTR)malloc(mSize);
		memmove((void*)aKey.buf, (aThisBuffer + 1 + sizeof(__int64)), mSize);
		if (aObject = BufferObject::Create((void*)aKey.buf, mSize))
			((BufferObject*)aObject)->SetBase(BufferObject::sPrototype);
	}
	else if (typeobj == -20 || typeobj == 20) //(_tcscmp(aObject->Type(), _T("ComValue")))
	{
		mSize = (size_t) * (__int64*)(aThisBuffer + 1);
		if (*(aThisBuffer + 1 + sizeof(__int64)) == VT_BSTR)
			aObject = new ComObject((long long)SysAllocString((OLECHAR*)(aThisBuffer + 2 + sizeof(__int64))), (VARTYPE) * (char*)(aThisBuffer + 1 + sizeof(__int64)), 0);
		else aObject = new ComObject(*(long long*)(aThisBuffer + 2 + sizeof(__int64)), (VARTYPE) * (char*)(aThisBuffer + 1 + sizeof(__int64)), 0);
	}
	else
		g_script->RuntimeError(ERR_TYPE_MISMATCH);
	if (!aObject)
		g_script->RuntimeError(ERR_OUTOFMEM);
	aObjects[aObjCount++] = aObject;


	aThisBuffer++;
	size_t aSize = (size_t) * (__int64*)aThisBuffer;
	aThisBuffer += sizeof(__int64);
	if (typeobj == -19 || typeobj == 19		//(_tcscmp(aObject->Type(), _T("Buffer")))
		|| typeobj == -20 || typeobj == 20)		//(_tcscmp(aObject->Type(), _T("ComValue")))
	{
		aThisBuffer += mSize;
		aSize -= mSize;
	}


	this_token.object = aObject;

	for (char* end = aThisBuffer + aSize; aThisBuffer < end;)
	{
		typekey = *(char*)aThisBuffer++;
		if (typekey < -12)
		{
			aThisBuffer -= 1;  // required for the object type
			aKey.symbol = SYM_OBJECT;
			aKey.object = ObjRawLoad(aThisBuffer, aObjects, aObjCount, aObjSize);
			aThisBuffer += 1 + sizeof(__int64) + *(__int64*)(aThisBuffer + 1);
		}
		else if (typekey == -12)
		{
			aKey.symbol = SYM_OBJECT;
			aKey.object = aObjects[*(__int64*)aThisBuffer];
			aKey.object->AddRef();
			aThisBuffer += sizeof(__int64);
		}
		else if (typekey == -10 || typekey == -11)
		{
			aKey.symbol = SYM_STRING;
			__int64 aMarkerSize = *(__int64*)aThisBuffer;
			aThisBuffer += sizeof(__int64);
			if (aMarkerSize)
			{
				aKey.marker_length = (size_t)(aMarkerSize - 1) / sizeof(TCHAR);
				aKey.marker = (LPTSTR)aThisBuffer;
				aThisBuffer += aMarkerSize;
			}
			else
			{
				aKey.marker = _T("");
				aKey.marker_length = 0;
			}
		}
		else if (typekey == -9)
		{
			aKey.symbol = SYM_STRING;
			aKey.marker_length = FTOA(*(double*)aThisBuffer, buf, MAX_INTEGER_LENGTH);
			aKey.marker = buf;
			aThisBuffer += sizeof(__int64);
		}
		else
		{
			aKey.symbol = SYM_INTEGER;
			if (typekey == -8)
			{
				aKey.value_int64 = *(__int64*)aThisBuffer;
				aThisBuffer += sizeof(__int64);
			}
			else if (typekey == -6)
			{
				aKey.value_int64 = *(UINT*)aThisBuffer;
				aThisBuffer += sizeof(UINT);
			}
			else if (typekey == -4)
			{
				aKey.value_int64 = *(USHORT*)aThisBuffer;
				aThisBuffer += sizeof(USHORT);
			}
			else if (typekey == -2)
			{
				aKey.value_int64 = *(BYTE*)aThisBuffer;
				aThisBuffer += sizeof(BYTE);
			}
			else if (typekey == -1)
			{
				aKey.value_int64 = *(char*)aThisBuffer;
				aThisBuffer += sizeof(char);
			}
			else if (typekey == -3)
			{
				aKey.value_int64 = *(short*)aThisBuffer;
				aThisBuffer += sizeof(short);
			}
			else if (typekey == -5)
			{
				aKey.value_int64 = *(int*)aThisBuffer;
				aThisBuffer += sizeof(int);
			}
			else if (typekey == -7)
			{
				aKey.value_int64 = *(__int64*)aThisBuffer;
				aThisBuffer += sizeof(__int64);
			}
			else
			{
				g_script->RuntimeError(ERR_INVALID_VALUE);
				return NULL;
			}
		}

		typeval = *(char*)aThisBuffer++;
		if (typeval > 12)
		{
			aValue.symbol = SYM_OBJECT;
			aValue.object = ObjRawLoad(--aThisBuffer, aObjects, aObjCount, aObjSize);   // aThisBuffer-- required to pass type of object
			aThisBuffer += 1 + sizeof(__int64) + *(__int64*)(aThisBuffer + 1);
		}
		else if (typeval == 12)
		{
			aValue.symbol = SYM_OBJECT;
			aValue.object = aObjects[*(__int64*)aThisBuffer];
			aValue.object->AddRef();
			aThisBuffer += sizeof(__int64);
		}
		else if (typeval == 10)
		{
			aValue.symbol = SYM_STRING;
			__int64 aMarkerSize = *(__int64*)aThisBuffer;
			aThisBuffer += sizeof(__int64);
			if (aMarkerSize)
			{
				aValue.marker_length = (size_t)(aMarkerSize - 1) / sizeof(TCHAR);
				aValue.marker = (LPTSTR)aThisBuffer;
				aThisBuffer += aMarkerSize;
			}
			else
			{
				aValue.marker = _T("");
				aValue.marker_length = 0;
			}
		}
		else if (typeval == 9)
		{
			aValue.symbol = SYM_FLOAT;
			aValue.value_double = *(double*)aThisBuffer;
			aThisBuffer += sizeof(__int64);
		}
		else
		{
			aValue.symbol = SYM_INTEGER;
			if (typeval == 8)
			{
				aValue.value_int64 = *(__int64*)aThisBuffer;
				aThisBuffer += sizeof(__int64);
			}
			else if (typeval == 6)
			{
				aValue.value_int64 = *(UINT*)aThisBuffer;
				aThisBuffer += sizeof(UINT);
			}
			else if (typeval == 4)
			{
				aValue.value_int64 = *(USHORT*)aThisBuffer;
				aThisBuffer += sizeof(USHORT);
			}
			else if (typeval == 2)
			{
				aValue.value_int64 = *(BYTE*)aThisBuffer;
				aThisBuffer += sizeof(BYTE);
			}
			else if (typeval == 1)
			{
				aValue.value_int64 = *(char*)aThisBuffer;
				aThisBuffer += sizeof(char);
			}
			else if (typeval == 3)
			{
				aValue.value_int64 = *(short*)aThisBuffer;
				aThisBuffer += sizeof(short);
			}
			else if (typeval == 5)
			{
				aValue.value_int64 = *(int*)aThisBuffer;
				aThisBuffer += sizeof(int);
			}
			else if (typeval == 7)
			{
				aValue.value_int64 = *(__int64*)aThisBuffer;
				aThisBuffer += sizeof(__int64);
			}
			else
			{
				g_script->RuntimeError(ERR_INVALID_VALUE);
				return NULL;
			}
		}
		result_token.InitResult(L"");
		if (typekey == -11)
			aObject->Invoke(result_token, IT_SET, aKey.marker, this_token, params + 1, 1);
		else if (typeobj == -15 || typeobj == 15 || typeobj == -16 || typeobj == 16)
			aObject->Invoke(result_token, IT_CALL, _T("Push"), this_token, params + 1, 1);
		else
			aObject->Invoke(result_token, IT_SET, 0, this_token, params, 2);
		aKey.Free();
		aValue.Free();
	}
	return aObject;
}

//
// ObjLoad()
//

BIF_DECL(BIF_ObjLoad)
{
	aResultToken.symbol = SYM_OBJECT;
	bool aFreeBuffer = false;
	DWORD aSize = aParamCount > 1 ? (DWORD)TokenToInt64(*aParam[1]) : 0;
	LPTSTR aPath = TokenToString(*aParam[0]);
	IObject* buffer_obj;
	char* aBuffer;
	size_t  max_bytes = SIZE_MAX;

	if (TokenIsNumeric(**aParam))
		aBuffer = (char*)TokenToInt64(**aParam);
	else if (buffer_obj = TokenToObject(**aParam))
	{
		size_t ptr;
		GetBufferObjectPtr(aResultToken, buffer_obj, ptr, max_bytes);
		if (aResultToken.Exited())
			return;
		aBuffer = (char*)ptr;
	}
	else
	{ // FileRead Mode
		if (GetFileAttributes(aPath) == 0xFFFFFFFF)
		{
			aResultToken.SetValue(_T(""));
			return;
		}
		FILE* fp;

		fp = _tfopen(aPath, _T("rb"));
		if (fp == NULL)
		{
			aResultToken.SetValue(_T(""));
			return;
		}

		fseek(fp, 0, SEEK_END);
		aSize = ftell(fp);
		if (!(aBuffer = (char*)malloc(aSize)))
		{
			aResultToken.SetValue(_T(""));
			fclose(fp);
			return;
		}
		aFreeBuffer = true;
		fseek(fp, 0, SEEK_SET);
		fread(aBuffer, 1, aSize, fp);
		fclose(fp);
	}
	if (*(unsigned int*)aBuffer == 0x04034b50)
	{
		LPVOID aDataBuf;
		aSize = *(ULONG*)((UINT_PTR)aBuffer + 8);
		if (*(ULONG*)((UINT_PTR)aBuffer + 16) > aSize)
			aSize = *(ULONG*)((UINT_PTR)aBuffer + 16);
		DWORD aSizeDeCompressed = DecompressBuffer(aBuffer, aDataBuf, aSize, ParamIndexIsOmittedOrEmpty(1) ? NULL : TokenToString(*aParam[1]));
		if (aSizeDeCompressed)
		{
			LPVOID buff = malloc(aSizeDeCompressed);
			aFreeBuffer = true;
			memcpy(buff, aDataBuf, aSizeDeCompressed);
			g_memset(aDataBuf, 0, aSizeDeCompressed);
			free(aDataBuf);
			aBuffer = (char*)buff;
		}
		else
			_o_throw(_T("ObjLoad: Password mismatch."));
	}
	UINT aObjCount = 0;
	UINT aObjSize = 16;
	IObject** aObjects = (IObject**)malloc(aObjSize * sizeof(IObject**));
	if (!aObjects || !(aResultToken.object = ObjRawLoad(aBuffer, aObjects, aObjCount, aObjSize)))
	{
		if (!TokenToInt64(*aParam[0]))
			free(aBuffer);
		free(aObjects);
		aResultToken.SetValue(_T(""));
		return;
	}
	free(aObjects);
	if (aFreeBuffer)
		free(aBuffer);
}



int ConvertDllArgTypes(LPTSTR& aBuf, DYNAPARM* aDynaParam, int aParamCount);
BIF_DECL(BIF_DynaCall)
{
	DynaToken* obj = new DynaToken();
	if (!obj)
		_o_throw_oom;
	obj->SetBase(DynaToken::sPrototype);

	IObjPtr dyobj{ obj };
	ExprTokenType token{}, miss(0);
	DYNAPARM* dyna_param = nullptr;
	int* param_shift = nullptr, i;
	char* param_free = nullptr;

	switch (TypeOfToken(*aParam[0]))
	{
	case SYM_STRING:
		obj->mFunction = (void*)GetDllProcAddress(ParamIndexToString(0));
		break;
	case SYM_INTEGER:
		obj->mFunction = (void*)ParamIndexToInt64(0);
		break;
	case SYM_OBJECT:
		size_t n;
		GetBufferObjectPtr(aResultToken, ParamIndexToObject(0), n);
		if (aResultToken.Exited())
			return;
		obj->mFunction = (void*)n;
		break;
	default:
		_o_throw(ERR_PARAM1_INVALID, ErrorPrototype::Type);
	}
	if (!obj->mFunction)
		_o_throw(ERR_NONEXISTENT_FUNCTION);
	else if (obj->mFunction < (void*)65536)
		_o_throw_param(0);

#ifdef WIN32_PLATFORM
	obj->mDllCallMode = DC_CALL_STD; // Set default.  Can be overridden to DC_CALL_CDECL and flags can be OR'd into it.
#endif
	obj->mReturnAttrib.type = DLL_ARG_INT;
	// Check validity of this arg's return type:
	Array* arr = nullptr;
	if (aParamCount > 1)
	{
		LPTSTR return_type_string;
		switch (TypeOfToken(*aParam[1]))
		{
		case SYM_STRING:
			return_type_string = ParamIndexToString(1);
			break;
		case SYM_OBJECT:
			if ((arr = dynamic_cast<Array*>(ParamIndexToObject(1))) && arr->ItemToToken(0, token) && token.symbol == SYM_STRING) {
				return_type_string = token.marker;
				break;
			}
		default:
			_o_throw(ERR_PARAM2_INVALID, ErrorPrototype::Type);
		}

		return_type_string = omit_leading_whitespace(return_type_string);

		// 64-bit note: The calling convention detection code is preserved here for script compatibility.
		LPTSTR equ;
		if (equ = _tcschr(return_type_string, '=')) {
			if (equ != return_type_string) {
				ConvertDllArgType(return_type_string, obj->mReturnAttrib, &i);
				if (return_type_string + i != equ && !IS_SPACE_OR_TAB(return_type_string[i]))
					_f_throw_value(ERR_INVALID_RETURN_TYPE, return_type_string);
			}
			return_type_string = omit_leading_whitespace(equ + (equ[1] == '=' || equ[1] == '@' ? 2 : 1));
#ifdef WIN32_PLATFORM
			if (equ[1] == '=')
				obj->mDllCallMode = DC_CALL_CDECL;
			else if (equ[1] == '@')
				obj->mDllCallMode = DC_CALL_THISCALL;

			if (!obj->mReturnAttrib.passed_by_address) // i.e. the special return flags below are not needed when an address is being returned.
			{
				if (obj->mReturnAttrib.type == DLL_ARG_DOUBLE)
					obj->mDllCallMode |= DC_RETVAL_MATH8;
				else if (obj->mReturnAttrib.type == DLL_ARG_FLOAT)
					obj->mDllCallMode |= DC_RETVAL_MATH4;
			}
#endif
		}

		for (i = 0; return_type_string[i]; i++)
			if (!_tcschr(_T("UP*0123456789 \t"), ctoupper(return_type_string[i])))
				obj->mParamCount++;

		obj->mData = malloc(obj->mParamCount * (sizeof(DYNAPARM) + sizeof(int) + sizeof(char)));
		dyna_param = (DYNAPARM*)obj->mData;
		if (!dyna_param || (i = ConvertDllArgTypes(return_type_string, dyna_param, obj->mParamCount)))
		{
			if (dyna_param)
				_o_throw_value(i > 0 ? ERR_PARAM_COUNT_INVALID : ERR_INVALID_ARG_TYPE, return_type_string);
			else _o_throw_oom;
		}
#ifdef WIN32_PLATFORM
		if (obj->mDllCallMode & DC_CALL_THISCALL)
			if (!obj->mParamCount)
				_f_throw_value(_T("Invalid calling convention."));
			else if (dyna_param[0].type != DLL_ARG_INT)
				_f_throw_value(ERR_INVALID_ARG_TYPE, return_type_string);
#endif
		param_shift = (int*)((char*)obj->mData + obj->mParamCount * sizeof(DYNAPARM));
		param_free = (char*)((char*)param_shift + obj->mParamCount * sizeof(int));
		memset(param_free, 0, obj->mParamCount);
	}

	ResultToken temp{};
	aParamCount -= 2, aParam += 2;
	for (i = 0; i < obj->mParamCount; i++)  // Same loop as used in DynaToken::Create below, so maintain them together.
	{
		DYNAPARM& this_dyna_param = dyna_param[i];
		ExprTokenType& this_param = i < aParamCount ? *aParam[i] : miss;
		IObject* this_param_obj = TokenToObject(this_param);
		if (this_param_obj)
		{
			if (this_dyna_param.is_ptr == 2) {
				this_dyna_param.ptr = this_param_obj;
				continue;
			}
			else
				// Support Buffer.Ptr, but only for "Ptr" type.  All other types are reserved for possible
				// future use, which might be general like obj.ToValue(), or might be specific to DllCall
				// or the particular type of this arg (Int, Float, etc.).
				if (this_dyna_param.is_ptr || this_dyna_param.type == DLL_ARG_STR || this_dyna_param.type == DLL_ARG_xSTR)
				{
					GetBufferObjectPtr(aResultToken, this_param_obj, this_dyna_param.value_uintptr);
					if (aResultToken.Exited())
						return;
					continue;
				}
		}

		SymbolType tp = TypeOfToken(this_param);
		switch (this_dyna_param.type)
		{
		case DLL_ARG_STR:
		case DLL_ARG_xSTR:
			if (tp == SYM_INTEGER)
				this_dyna_param.ptr = (void*)TokenToInt64(this_param);
			else if (tp == SYM_STRING) {
				size_t len;
				this_dyna_param.ptr = TokenToString(this_param, NULL, &len);
				param_free[i] = 1;
				if (len) {
					if (this_dyna_param.type == DLL_ARG_xSTR) {
#ifdef UNICODE
						this_dyna_param.ptr = CStringCharFromWChar((LPTSTR)this_dyna_param.ptr).DetachBuffer();
#else
						this_dyna_param.ptr = CStringWCharFromChar((LPTSTR)this_dyna_param.ptr).DetachBuffer();
#endif // UNICODE
					}
					else
						this_dyna_param.ptr = temp.Malloc((LPTSTR)this_dyna_param.ptr, len);
				}
				else this_dyna_param.ptr = L"", param_free[i] = 2;
			}
			else
				_o_throw_type(_T("String"), this_param);
			break;

		case DLL_ARG_DOUBLE:
		case DLL_ARG_FLOAT:
			// This currently doesn't validate that this_dyna_param.is_unsigned==false, since it seems
			// too rare and mostly harmless to worry about something like "Ufloat" having been specified.
			if (!TokenIsNumeric(this_param))
				_f_throw_type(_T("Number"), this_param);
			this_dyna_param.value_double = TokenToDouble(this_param);
			if (this_dyna_param.type == DLL_ARG_FLOAT)
				this_dyna_param.value_float = (float)this_dyna_param.value_double;
			break;

		case DLL_ARG_STRUCT: {
			int& size = this_dyna_param.struct_size;
			size = (size + sizeof(void*) - 1) & -(int)sizeof(void*);
			if (this_param_obj) {
				GetBufferObjectPtr(aResultToken, this_param_obj, this_dyna_param.value_uintptr);
				if (aResultToken.Exited())
					return;
			}
			else if (TokenIsPureNumeric(this_param) == SYM_INTEGER)
				this_dyna_param.value_int64 = TokenToInt64(this_param);
			else
				_o_throw_type(_T("Integer"), this_param);
			if (this_dyna_param.value_uintptr < 65536 && (&this_param != &miss))
				_o_throw_value(ERR_INVALID_VALUE);
			break;
		}
		default: // Namely:
		//case DLL_ARG_INT:
		//case DLL_ARG_SHORT:
		//case DLL_ARG_CHAR:
		//case DLL_ARG_INT64:
			if (tp == SYM_STRING && this_dyna_param.is_ptr == 1) {
				this_dyna_param.ptr = TokenToString(this_param);
				if (this_param.symbol == SYM_VAR)
					this_dyna_param.ptr = _tcsdup((LPTSTR)this_dyna_param.ptr), param_free[i] = true;
			}
			else if (!TokenIsNumeric(this_param))
				_f_throw_type(_T("Number"), this_param);
			else this_dyna_param.value_int64 = TokenToInt64(this_param);
		} // switch (this_dyna_param.type)
	} // for() each arg.

	for (i = 0; i < obj->mParamCount; i++)
		param_shift[i] = i;
	UINT arrLen;
	if (arr && (arrLen = arr->Length()) > 1)
	{
		// Set shift info for parameters, -1 for definition in first item.
		memset(param_shift, -1, obj->mParamCount * sizeof(int));
		bool* paramslist = (bool*)_alloca(obj->mParamCount);
		memset(paramslist, 0, obj->mParamCount);
		for (UINT i = 1; i < arrLen; i++) {
			arr->ItemToToken(i, token);
			if (token.symbol != SYM_INTEGER)
				_o_throw_type(_T("Integer"), token);
			else if (token.value_int64 < 1 || token.value_int64 > obj->mParamCount || paramslist[token.value_int64 - 1])
				_o_throw(ERR_INVALID_INDEX, TokenToString(token, _f_number_buf), ErrorPrototype::Index);
			param_shift[i] = (int)(token.value_int64 - 1);
			paramslist[token.value_int64 - 1] = true;
		}
		for (i = 0; i < obj->mParamCount; i++) {
			if (paramslist[i]) continue;
			for (int j = 0; j < obj->mParamCount; j++)
				if (param_shift[j] == -1) {
					param_shift[j] = i;
					goto next;
				}
		next:;
		}
	}
	dyobj.obj = nullptr;
	_o_return(obj);
}



//
// DynaToken::Create - Called by BIF_DynaCall to create a new object, optionally passing key/value pairs to set.
//

DynaToken* DynaToken::Create(ExprTokenType* aParam[], int aParamCount) {
	ResultToken result{};
	TCHAR buf[MAX_NUMBER_SIZE];
	result.InitResult(buf);
	BIF_DynaCall(result, aParam, aParamCount);
	return result.symbol == SYM_OBJECT ? (DynaToken*)result.object : nullptr;
}



DynaToken::~DynaToken()
{
	if (mData) {
		char* param_free = ((char*)mData + mParamCount * (sizeof(DYNAPARM) + sizeof(int)));
		for (int i = 0; i < mParamCount; i++)
			if (param_free[i] == 1)
				free(((DYNAPARM*)mData)[i].ptr);
		free(mData);
	}
}



void DynaToken::Invoke(ResultToken& aResultToken, int aID, int aFlags, ExprTokenType* aParam[], int aParamCount) {
	switch (aID)
	{
	case P_MinParams:
	case P_MaxParams:
		_o_return(aID == P_MinParams ? 0 : mParamCount);

	case P_Param: {
		DYNAPARM* dyna_param = (DYNAPARM*)mData;
		int* param_shift = (int*)((char*)mData + mParamCount * sizeof(DYNAPARM));
		ExprTokenType* val;
		SymbolType symbol;
		if (IS_INVOKE_SET)
			val = aParam[0], symbol = TypeOfToken(*val), aParam++, aParamCount--;
		auto index = ParamIndexToInt64(0);
		if (index < 1 || index > mParamCount)
			_o_throw(ERR_INVALID_INDEX, ParamIndexToString(0, _f_number_buf), ErrorPrototype::Index);
		auto& param = dyna_param[index = param_shift[--index]];
		char& f = ((char*)((char*)mData + mParamCount * (sizeof(DYNAPARM) + sizeof(int))))[index];
		switch (param.type)
		{
		case DLL_ARG_STR:
		case DLL_ARG_xSTR:
			if (IS_INVOKE_GET) {
				if (f) _o_return((LPTSTR)param.ptr);
				else _o_return(param.value_int64);
			}
			else {
				UINT_PTR ptr;
				size_t len;
				if (symbol == SYM_OBJECT) {
					GetBufferObjectPtr(aResultToken, TokenToObject(*val), ptr);
					if (aResultToken.Exited())
						return;
				}
				else if (symbol == SYM_INTEGER)
					ptr = (UINT_PTR)TokenToInt64(*val);
				else if (symbol == SYM_STRING)
					ptr = (UINT_PTR)TokenToString(*val, NULL, &len);
				else _o_throw_type(_T("String"), *val);
				if (f == 1) free(param.ptr);
				param.ptr = (void*)ptr, f = 0;;
				if (symbol == SYM_STRING) {
					f = 1;
					if (len) {
						if (param.type == DLL_ARG_xSTR) {
#ifdef UNICODE
							param.ptr = CStringCharFromWChar((LPTSTR)param.ptr).DetachBuffer();
#else
							param.ptr = CStringWCharFromChar((LPTSTR)param.ptr).DetachBuffer();
#endif // UNICODE
						}
						else
							param.ptr = ResultToken().Malloc((LPTSTR)param.ptr, len);
					}
					else param.ptr = L"", f = 2;
				}
			}
			break;

		case DLL_ARG_FLOAT:
			if (IS_INVOKE_GET)
				_o_return(param.value_float);
			else if (TokenIsNumeric(*val))
				param.value_float = (float)TokenToDouble(*val);
			else
				_o_throw_type(_T("Number"), *val);
			break;

		case DLL_ARG_DOUBLE:
			if (IS_INVOKE_GET)
				_o_return(param.value_double);
			else if (TokenIsNumeric(*val))
				param.value_double = TokenToDouble(*val);
			else
				_o_throw_type(_T("Number"), *val);
			break;

		case DLL_ARG_STRUCT:
			if (IS_INVOKE_GET)
				_o_return(param.value_int64);
			else {
				UINT_PTR ptr;
				if (symbol == SYM_OBJECT) {
					GetBufferObjectPtr(aResultToken, TokenToObject(*val), ptr);
					if (aResultToken.Exited())
						return;
				}
				else if (symbol == SYM_INTEGER)
					ptr = (UINT_PTR)TokenToInt64(*val);
				else _o_throw_type(_T("Integer"), *val);
				if (ptr < (UINT_PTR)65536)
					_o_throw_value(ERR_PARAM_INVALID, TokenToString(*val, aResultToken.buf));
				param.value_uintptr = ptr;
			}
			break;

		default:
			if (IS_INVOKE_GET) {
				if (param.type == DLL_ARG_INT64)
					_o_return(param.value_int64);
				if (param.type == DLL_ARG_INT)
					_o_return(param.is_unsigned ? (UINT)param.value_int : param.value_int);
				if (param.type == DLL_ARG_SHORT)
					_o_return(param.is_unsigned ? param.value_int & 0x0000FFFF : (SHORT)(WORD)param.value_int);
				if (param.type == DLL_ARG_CHAR)
					_o_return(param.is_unsigned ? param.value_int & 0x000000FF : (char)(BYTE)param.value_int);
			}
			else if (IS_NUMERIC(symbol))
				param.value_int64 = TokenToInt64(*val);
			else
				_o_throw_type(_T("Number"), *val);
			break;
		}
		break;
	}
	default:
		break;
	}
}



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
	double_conversion::StringToDoubleConverter converter(4, 0, __builtin_nan("0"), 0, 0);
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
	if (b < (void*)65536)
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
		if (beg > (void*)65535)
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
					val.SetValue(ATOF(t, &beg));
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
		if (obj && field.prop->Getter() && field.prop->MaxParams <= 0 && !((Object *)obj)->IsClassPrototype()) {
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
		IndexEnumerator* enumobj = dynamic_cast<IndexEnumerator*>(enumerator);
		if (enumobj || dynamic_cast<EnumBase*>(enumerator)) {
			char enumtype = enumobj ? 1 : 0;
			if (enumobj && dynamic_cast<Array*>(enumobj->mObject))
				enumobj->mParamCount = 2, enumtype++;
			if (enumtype == 1) {
				str.append('{'), ++deep;
				while (((EnumBase*)enumerator)->Next(&vkey, &vval) == CONDITION_TRUE)
					append(vkey, vval);
				vkey.Free(), vval.Free();
				if (str.back() == ',')
					str.pop_back(), append(--deep), str.append('}');
				else
					--deep, str.append('}');
			}
			else if (enumtype == 2 || dynamic_cast<ComArrayEnum*>(enumerator)) {
				if (!enumobj)
					((ComArrayEnum*)enumerator)->mIndexMode = true;
				str.append('['), ++deep;
				while (((EnumBase*)enumerator)->Next(&vkey, &vval) == CONDITION_TRUE)
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
				if (((ComEnum*)enumerator)->Next(&vkey, nullptr) == CONDITION_TRUE) {
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
						while (((ComEnum*)enumerator)->Next(&vkey, nullptr) == CONDITION_TRUE) {
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
						while (((ComEnum*)enumerator)->Next(&vkey, nullptr) == CONDITION_TRUE)
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


void Promise::Init(int aParamCount) {
	if (!aParamCount)
		return;
	if (mParamToken = (ResultToken*)malloc(aParamCount * (sizeof(ResultToken*) + sizeof(ResultToken)))) {
		mParamCount = mMaxParamCount = aParamCount;
		mParam = (ResultToken**)(mParamToken + mParamCount);
		for (int i = 0; i < mParamCount; i++)
			mParam[i] = &mParamToken[i], mParamToken[i].mem_to_free = nullptr;
	}
}

void Promise::Init(ExprTokenType* aParam[], int aParamCount) {
	if (!aParamCount)
		return;
	if (mParamToken = (ResultToken*)malloc(aParamCount * (sizeof(ResultToken*) + sizeof(ResultToken)))) {
		mParamCount = mMaxParamCount = aParamCount;
		mParam = (ResultToken**)(mParamToken + mParamCount);
		for (int i = 0; i < mParamCount; i++) {
			mParam[i] = &mParamToken[i];
			if (aParam[i]->symbol == SYM_STRING)
				mParamToken[i].Malloc(aParam[i]->marker, aParam[i]->marker_length);
			else {
				mParamToken[i].mem_to_free = nullptr;
				mParamToken[i].CopyValueFrom(*aParam[i]);
				if (mParamToken[i].symbol == SYM_OBJECT) {
					if (mMarshal)
						MarshalObjectToToken(mParamToken[i].object, mParamToken[i]);
					else
						mParamToken[i].object->AddRef();
				}
			}
		}
	}
}

void callPromise(Promise* aPromise, HWND replyHwnd);
void Promise::Invoke(ResultToken& aResultToken, int aID, int aFlags, ExprTokenType* aParam[], int aParamCount)
{
	Object *callback = dynamic_cast<Object *>(TokenToObject(*aParam[0]));
	if (!callback)
		_o_throw(ERR_INVALID_FUNCTOR);
	if (!ValidateFunctor(callback, 1, aResultToken))
		return;
	EnterCriticalSection(&mCritical);
	callback->AddRef();
	if (aID == M_Then) {
		if (mComplete)
			mComplete->Release();
		mComplete = callback;
		if (!(mState & 1)) callback = nullptr;
	}
	else {
		if (mError)
			mError->Release();
		mError = callback;
		if (!(mState & 2)) callback = nullptr;
	}
	LeaveCriticalSection(&mCritical);
	if (callback) {
		AddRef();
		callPromise(this, (HWND)-1);
	}
	AddRef();
	_f_return(this);
}

void Promise::FreeParams() {
	LPSTREAM stream;
	for (int i = 0; i < mParamCount; i++) {
		if (mParam[i]->symbol == SYM_STREAM) {
			if (stream = (LPSTREAM)mParam[i]->value_int64)
				CoReleaseMarshalData(stream), stream->Release();
		}
		else mParam[i]->Free();
	}
	mParamCount = 0;
}

void Promise::FreeResult() {
	LPSTREAM stream;
	if (mResult.symbol == SYM_STREAM) {
		if (stream = (LPSTREAM)mResult.value_int64)
			CoReleaseMarshalData(stream), stream->Release();
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

Func* Promise::ToBoundFunc() {
	if (!sCaller)
		sCaller = new BuiltInFunc(_T(""), [](BIF_DECL_PARAMS) {
		auto promise = (Promise*)aParam[0]->object;
		callPromise(promise, promise->mReply);
			}, 1, 1);
	ExprTokenType t_this(this), *param[] = { &t_this };
	return BoundFunc::Bind(sCaller, IT_CALL, nullptr, param, 1);
}

Worker* Worker::Create()
{
	auto obj = new Worker();
	obj->SetBase(sPrototype);
	return obj;
}

bool Worker::New(DWORD aThreadID)
{
	if (aThreadID == 0)
		mIndex = 0;
	else {
		mIndex = -1;
		for (int i = 0; i < MAX_AHK_THREADS; i++)
			if (g_ahkThreads[i].ThreadID == aThreadID)
			{
				mIndex = i;
				break;
			}
		if (mIndex == -1 && aThreadID < MAX_AHK_THREADS && g_ahkThreads[aThreadID].Hwnd)
			mIndex = aThreadID;
	}
	if (mIndex != -1 && (mThread = OpenThread(THREAD_ALL_ACCESS, TRUE, mThreadID = g_ahkThreads[mIndex].ThreadID))) {
		mHwnd = g_ahkThreads[mIndex].Hwnd;
		return true;
	}
	return false;
}

bool Worker::New(LPTSTR aScript, LPTSTR aCmd, LPTSTR aTitle)
{
	if (DWORD threadid = NewThread(aScript, aCmd, aTitle))
		return Worker::New(threadid);
	return false;
}

void Worker::Invoke(ResultToken& aResultToken, int aID, int aFlags, ExprTokenType* aParam[], int aParamCount)
{
	if (aID == M___New) {
		if (ParamIndexIsNumeric(0)) {
			if (New((DWORD)TokenToInt64(*aParam[0])))
				_o_return_retval;
			_f_throw(_T("Invalid ThreadID"));
		}
		else {
			if (New(ParamIndexToString(0), ParamIndexToOptionalString(1), ParamIndexToOptionalStringDef(2, _T("AutoHotkey"))))
				_o_return_retval;
			_f_return_FAIL;
		}
	}
	if (!mThread || !mHwnd || g_ahkThreads[mIndex].Hwnd != mHwnd) {
		mHwnd = NULL;
		if (mThread && WaitForSingleObject(mThread, 500) == WAIT_TIMEOUT) {
			for (int i = 0; i < MAX_AHK_THREADS; i++) {
				if (g_ahkThreads[i].ThreadID == mThreadID) {
					mIndex = i, mHwnd = g_ahkThreads[i].Hwnd;
					break;
				}
			}
		}
		if (!mHwnd) {
			if (aID == M_Wait || aID == M_ExitApp)
				_f_return_retval;
			if (aID == P_Ready)
				_f_return(0);
			_f_throw_win32(RPC_E_SERVER_DIED_DNE);
		}
	}

	ExprTokenType* val = NULL;
	LPSTREAM stream = NULL;
	IUnknown* obj = NULL;
	AutoTLS atls;
	Script* script = g_ahkThreads[mIndex].Script;
	Var* var = nullptr;
	LPTSTR name;
#ifdef _WIN64
	DWORD ThreadID = __readgsdword(0x48);
#else
	DWORD ThreadID = __readfsdword(0x24);
#endif

	switch (aID)
	{
	case P___Item:
	{
		if (IS_INVOKE_SET)
			val = aParam[0], aParam++, aParamCount--;
		if (mThreadID != ThreadID)
			atls.Lock(mIndex);
		var = g_ahkThreads[mIndex].Script->FindGlobalVar(name = TokenToString(*aParam[0]));
		ComObject* comobj = nullptr;
		VARIANT vv = { 0 };

		if (!var) {
			atls.~AutoTLS();
			_f_throw(ERR_DYNAMIC_NOT_FOUND, name);
		}
		bool iunk = false;
		if (IS_INVOKE_SET) {
			if (atls.tls) {
				ResultToken result;
				atls.~AutoTLS();
				if (obj = TokenToObject(*val)) {
					MarshalObjectToToken((IObject*)obj, result);
					val = &result;
				}
				else result.symbol = SYM_INTEGER;
				SendMessage(mHwnd, AHK_THREADVAR, (WPARAM)val, (LPARAM)var);
				if (result.symbol == SYM_STREAM) {
					stream = (LPSTREAM)result.value_int64;
					CoReleaseMarshalData(stream), stream->Release();
				}
			}
			else
				var->Assign(*val);
		}
		else {
			if (var->mType == VAR_VIRTUAL) {
				aResultToken.symbol = SYM_INTEGER;
				var->Get(aResultToken);
				if (aResultToken.symbol == SYM_OBJECT)
					aResultToken.object->AddRef();
			}
			else
				var->ToToken(aResultToken);
			if (atls.tls) {
				if (aResultToken.symbol == SYM_STRING)
					!aResultToken.mem_to_free &&aResultToken.Malloc(aResultToken.marker, aResultToken.marker_length ? aResultToken.marker_length : -1);
				else if (obj = TokenToObject(aResultToken)) {
					IObjPtr to_free{ (IObject *)obj };
					if (comobj = dynamic_cast<ComObject *>(obj)) {
						if (comobj->mVarType == VT_DISPATCH || comobj->mVarType == VT_UNKNOWN)
							obj = comobj->mUnknown, iunk = comobj->mVarType == VT_UNKNOWN;
						else if (!(comobj->mVarType & ~VT_TYPEMASK)) {
							comobj->ToVariant(vv);
							VariantToToken(vv, aResultToken);
							obj = nullptr;
						}
					}
					if (obj) {
						atls.~AutoTLS();
						IObject *pobj;
						if ((stream = (LPSTREAM)SendMessage(mHwnd, AHK_THREADVAR, (WPARAM)obj, !iunk)) && (pobj = UnMarshalObjectFromStream(stream)))
							_f_return(pobj);
						_f_throw_win32();
					}
				}
			}
		}
		_f_return_retval;
	}
	case M_AddScript:
		_f_return(addScript(ParamIndexToString(0), ParamIndexToOptionalInt(1, 0), mThreadID));
	case M_AsyncCall:
	{
		Object* func = nullptr;
		Var* var = nullptr;
		if (mThreadID != ThreadID)
			atls.Lock(mIndex);
		else
			func = dynamic_cast<Object*>(TokenToObject(*aParam[0]));
		if (!func) {
			if (*(name = TokenToString(*aParam[0])))
				var = g_ahkThreads[mIndex].Script->FindGlobalVar(name);
			if (!var) {
				atls.~AutoTLS();
				_f_throw(ERR_DYNAMIC_NOT_FOUND, name);
			}
			func = dynamic_cast<Object*>(var->ToObject());
		}
		if (!func || !func->HasMethod(_T("Call"))) {
			atls.~AutoTLS();
			ExprTokenType func_token;
			if (var)
				func_token.SetVar(var);
			else func_token.SetValue(func);
			aResultToken.UnknownMemberError(func_token, IT_CALL, nullptr);
			return;
		}
		aParam++, aParamCount--;
		atls.~AutoTLS();
		auto promise = new Promise(func, !!atls.teb);
		promise->Init(aParam, aParamCount);
		promise->mPriority = g->Priority;
		if (!SendMessage(mHwnd, AHK_EXECUTE_FUNCTION, (WPARAM)g_hWnd, (LPARAM)promise)) {
			auto err = GetLastError();
			promise->Release();
			_f_throw_win32(err);
		}
		promise->AddRef();
		_f_return(promise);
	}
	case M_Exec:
		_f_return(ahkExec(ParamIndexToString(0), atls.tls ? mThreadID : ParamIndexToOptionalBOOL(1, TRUE) ? 0 : mThreadID));
	case M_ExitApp:
		if (script)
			TerminateSubThread(mThreadID, mHwnd);
		CloseHandle(mThread), mThread = NULL;
		_f_return_retval;
	case M_Pause:
		_f_return(ahkPause(ParamIndexToBOOL(0), mThreadID));
	case M_Reload:
		if (!PostMessage(mHwnd, WM_COMMAND, ID_FILE_RELOADSCRIPT, 0))
			_f_return(0);
		_f_return(1);
	case P_Ready:
		_f_return(script ? script->mIsReadyToExecute : 0);
	case P_ThreadID:
		_f_return(mThreadID);
	case M_Wait:
		if (mThreadID == ThreadID)
			_f_return(0);
		auto timeout = ParamIndexToOptionalInt64(0, 0);
		auto g = ::g;
		for (int start_time = GetTickCount(); WaitForSingleObject(mThread, 0) == WAIT_TIMEOUT;) {
			if ((char)g->IsPaused == -1)	// thqby: Used to terminate a thread
				_f_return_retval;
			if (g_MainThreadID != ThreadID)
				Sleep(SLEEP_INTERVAL);
			else MsgSleep(INTERVAL_UNSPECIFIED);
			if (timeout && (int)(timeout - (GetTickCount() - start_time)) <= SLEEP_INTERVAL_HALF)
				_f_return(0);
		}
		_f_return(1);
	}
}

bool MarshalObjectToToken(IObject* aObj, ResultToken& aToken) {
	bool idisp = true;
	IUnknown* marsha = aObj;
	LPSTREAM stream;
	aToken.mem_to_free = nullptr;
	if (auto comobj = dynamic_cast<ComObject*>(aObj)) {
		if (comobj->mVarType == VT_DISPATCH || comobj->mVarType == VT_UNKNOWN)
			marsha = comobj->mUnknown, idisp = comobj->mVarType == VT_DISPATCH;
		else if (!(comobj->mVarType & ~VT_TYPEMASK)) {
			VARIANT vv = { 0 };
			comobj->ToVariant(vv);
			VariantToToken(vv, aToken);
			marsha = nullptr;
		}
	}
	if (!marsha)
		return false;
	if (SUCCEEDED(CoMarshalInterThreadInterfaceInStream(idisp ? IID_IDispatch : IID_IUnknown, marsha, &stream)) ||
		(idisp && SUCCEEDED(CoMarshalInterThreadInterfaceInStream(IID_IUnknown, marsha, &stream))))
		aToken.value_int64 = (__int64)stream, aToken.symbol = SYM_STREAM;
	else aObj->AddRef(), aToken.SetValue(aObj);
	return true;
}