#ifndef _LITEUNZIP_H
#define _LITEUNZIP_H

/*
* LiteUnzip.h
*
* For decompressing the contents of zip archives using LITEUNZIP.DLL.
*
* This file is a repackaged form of extracts from the zlib code available
* at www.gzip.org/zlib, by Jean-Loup Gailly and Mark Adler. The original
* copyright notice may be found in unzip.cpp. The repackaging was done
* by Lucian Wischik to simplify and extend its use in Windows/C++. Also
* encryption and unicode filenames have been added. Code was further
* revamped and turned into a DLL by Jeff Glatt.
*/

#ifdef __cplusplus
extern "C" {
#endif

#include <windows.h>

	// An HUNZIP identifies a zip archive that has been opened
#define HUNZIP	void *

	// Struct used to retrieve info about an entry in an archive
	typedef struct
	{
		ULONGLONG		Index;					// index of this entry within the zip archive
		DWORD			Attributes;				// attributes, as in GetFileAttributes.
		FILETIME		AccessTime, CreateTime, ModifyTime;	// access, create, modify filetimes
		ULONGLONG		CompressedSize;			// sizes of entry, compressed and uncompressed. These
		ULONGLONG		UncompressedSize;		// may be -1 if not yet known (e.g. being streamed in)
		ULONGLONG		offset;
		WCHAR			Name[MAX_PATH];			// entry name
	} ZIPENTRY;

	// Functions for opening a ZIP archive
	DWORD WINAPI UnzipOpenFileW(HUNZIP *, const WCHAR *, const char *, UINT);
	DWORD WINAPI UnzipOpenFileRawW(HUNZIP *, const WCHAR *, const char *);
#define UnzipOpenFile UnzipOpenFileW
#define UnzipOpenFileRaw UnzipOpenFileRawW
	DWORD WINAPI UnzipOpenBuffer(HUNZIP *, void *, ULONGLONG, const char *, UINT);
	DWORD WINAPI UnzipOpenBufferRaw(HUNZIP *, void *, ULONGLONG, const char *);
	DWORD WINAPI UnzipOpenHandle(HUNZIP *, HANDLE, const char *, UINT);
	DWORD WINAPI UnzipOpenHandleRaw(HUNZIP *, HANDLE, const char *);

	// Functions to get information about an "entry" within a ZIP archive
	DWORD WINAPI UnzipGetItemW(HUNZIP, ZIPENTRY *);
#define UnzipGetItem UnzipGetItemW
	DWORD WINAPI UnzipFindItemW(HUNZIP, ZIPENTRY *, BOOL);
#define UnzipFindItem UnzipFindItemW

	// Functions to unzip an "entry" within a ZIP archive
	DWORD WINAPI UnzipItemToFileW(HUNZIP, const WCHAR *, ZIPENTRY *);
#define UnzipItemToFile UnzipItemToFileW

	DWORD WINAPI UnzipItemToHandle(HUNZIP, HANDLE, ZIPENTRY *);
	DWORD WINAPI UnzipItemToBuffer(HUNZIP, void *, ZIPENTRY *);

	// Function to set the base directory
	DWORD WINAPI UnzipSetBaseDirW(HUNZIP, const WCHAR *);
#define UnzipSetBaseDir UnzipSetBaseDirW

	// Function to close an archive
	DWORD WINAPI UnzipClose(HUNZIP);

#define UnzipFormatMessage ZipFormatMessageW

#if !defined(ZR_OK)
	// These are the return codes from Unzip functions
#define ZR_OK			0		// Success
	// The following come from general system stuff (e.g. files not openable)
#define ZR_NOFILE		1		// Can't create/open the file
#define ZR_NOALLOC		2		// Failed to allocate memory
#define ZR_WRITE		3		// A general error writing to the file
#define ZR_NOTFOUND		4		// Can't find the specified file in the zip
#define ZR_MORE			5		// There's still more data to be unzipped
#define ZR_CORRUPT		6		// The zipfile is corrupt or not a zipfile
#define ZR_READ			7		// An error reading the file
#define ZR_NOTSUPPORTED	8		// The entry is in a format that can't be decompressed by this Unzip add-on
	// The following come from mistakes on the part of the caller
#define ZR_ARGS			9		// Bad arguments passed
#define ZR_NOTMMAP		10		// Tried to ZipGetMemory, but that only works on mmap zipfiles, which yours wasn't
#define ZR_MEMSIZE		11		// The memory-buffer size is too small
#define ZR_FAILED		12		// Already failed when you called this function
#define ZR_ENDED		13		// The zip creation has already been closed
#define ZR_MISSIZE		14		// The source file size turned out mistaken
#define ZR_ZMODE		15		// Tried to mix creating/opening a zip 
	// The following come from bugs within the zip library itself
#define ZR_SEEK			16		// trying to seek in an unseekable file
#define ZR_NOCHANGE		17		// changed its mind on storage, but not allowed
#define ZR_FLATE		18		// An error in the de/inflation code
#define ZR_PASSWORD		19		// Password is incorrect
#define ZR_ABORT		20		// Zipping aborted
#endif

#ifdef __cplusplus
}
#endif

#endif // _LITEUNZIP_H


#ifndef _LITEZIP_H
#define _LITEZIP_H

#ifdef __cplusplus
extern "C" {
#endif

#include <windows.h>

	/*
	* ZIP.H
	*
	* For creating zip files using LITEZIP.DLL.
	*
	* This file is a repackaged form of extracts from the zlib code available
	* at www.gzip.org/zlib, by Jean-Loup Gailly and Mark Adler. The original
	* copyright notice may be found in LiteZip.c. The repackaging was done
	* by Lucian Wischik to simplify and extend its use in Windows/C++. Also
	* encryption and unicode filenames have been added. Code was further
	* revamped and turned into a DLL (which supports both UNICODE and ANSI
	* C strings) by Jeff Glatt. Linux version by Jeff Glatt.
	*/

	// An HZIP identifies a zip archive that is being created
#define HZIP	void *

	// Functions to create a ZIP archive
	DWORD WINAPI ZipCreateFileW(HZIP *, const WCHAR *, const char *, int);
#define ZipCreateFile ZipCreateFileW
	DWORD WINAPI ZipCreateBuffer(HZIP *, void *, ULONGLONG, const char *, int);
	DWORD WINAPI ZipCreateHandle(HZIP *, HANDLE, const char *, int);

	// Functions for adding a "file" to a ZIP archive
	DWORD WINAPI ZipAddFileW(HZIP, const WCHAR *, const WCHAR *);
	DWORD WINAPI ZipAddFileRawW(HZIP, const WCHAR *);
#define ZipAddFile ZipAddFileW
#define ZipAddFileRaw ZipAddFileRawW

	DWORD WINAPI ZipAddHandleW(HZIP, const WCHAR *, HANDLE);
#define ZipAddHandle ZipAddHandleW
	DWORD WINAPI ZipAddHandleRaw(HZIP, HANDLE);

	DWORD WINAPI ZipAddPipeW(HZIP, const WCHAR *, HANDLE, ULONGLONG);
#define ZipAddPipe ZipAddPipeW
	DWORD WINAPI ZipAddPipeRaw(HZIP, HANDLE, ULONGLONG);

	DWORD WINAPI ZipAddBufferW(HZIP, const WCHAR *, const void *, ULONGLONG);
#define ZipAddBuffer ZipAddBufferW
	DWORD WINAPI ZipAddBufferRaw(HZIP, const void *, ULONGLONG);

	DWORD WINAPI ZipAddFolderW(HZIP, const WCHAR *);
#define ZipAddFolder ZipAddFolderW

	// Function to get a pointer to the ZIP archive created in memory by ZipCreateBuffer(0, len)
	DWORD WINAPI ZipGetMemory(HZIP, void **, ULONGLONG *, HANDLE *);

	// Function to reset a memory-buffer ZIP archive to compress a new ZIP archive
	DWORD WINAPI ZipResetMemory(HZIP);

	// Function to close the ZIP archive
	DWORD WINAPI ZipClose(HZIP);

	// Function to set options for ZipAdd* functions
	DWORD WINAPI ZipOptions(HZIP, DWORD);

#define TZIP_OPTION_GZIP	0x80000000
#define TZIP_OPTION_ABORT	0x40000000

	// Function to get an appropriate error message for a given error code return by Zip functions
	DWORD WINAPI ZipFormatMessageW(DWORD, WCHAR *, DWORD);
#define ZipFormatMessage ZipFormatMessageW

#if !defined(ZR_OK)
	// These are the return codes from Zip functions
#define ZR_OK			0		// Success
	// The following come from general system stuff (e.g. files not openable)
#define ZR_NOFILE		1		// Can't create/open the file
#define ZR_NOALLOC		2		// Failed to allocate memory
#define ZR_WRITE		3		// A general error writing to the file
#define ZR_NOTFOUND		4		// Can't find the specified file in the zip
#define ZR_MORE			5		// There's still more data to be unzipped
#define ZR_CORRUPT		6		// The zipfile is corrupt or not a zipfile
#define ZR_READ			7		// An error reading the file
#define ZR_NOTSUPPORTED	8		// The entry is in a format that can't be decompressed by this Unzip add-on
	// The following come from mistakes on the part of the caller
#define ZR_ARGS			9		// Bad arguments passed
#define ZR_NOTMMAP		10		// Tried to ZipGetMemory, but that only works on mmap zipfiles, which yours wasn't
#define ZR_MEMSIZE		11		// The memory-buffer size is too small
#define ZR_FAILED		12		// Already failed when you called this function
#define ZR_ENDED		13		// The zip creation has already been closed
#define ZR_MISSIZE		14		// The source file size turned out mistaken
#define ZR_ZMODE		15		// Tried to mix creating/opening a zip 
	// The following come from bugs within the zip library itself
#define ZR_SEEK			16		// trying to seek in an unseekable file
#define ZR_NOCHANGE		17		// changed its mind on storage, but not allowed
#define ZR_FLATE		18		// An error in the de/inflation code
#define ZR_PASSWORD		19		// Password is incorrect
#define ZR_ABORT		20		// Zipping aborted
#endif

#ifdef __cplusplus
}
#endif

typedef unsigned char	UCH;
typedef unsigned long	ULG;
ULG crc32(ULG crc, const unsigned char* buf, DWORD len);

#endif // _LITEZIP_H
