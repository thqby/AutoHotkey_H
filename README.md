# AutoHotkey_H v2.1 #

- [Changes](#changes-from-ahkdll)
- [Classes List](#classes-list)
- [Functions List](#functions-list)
- [AutoHotkey Module](#autohotkey.dll-module)
  - [COM Interfaces](#com-interfaces)
  - [Export Functions](#export-functions)
- [Compile](#how-to-compile)

[AutoHotkey](https://autohotkey.com/) is a free, open source macro-creation and automation software utility that allows users to automate repetitive tasks. It is driven by a custom scripting language that is aimed specifically at providing keyboard shortcuts, otherwise known as hotkeys.

AutoHotkey_H v2 started as a fork of [AutoHotkey_L v2](https://github.com/Lexikos/AutoHotkey_L/tree/alpha), merge branch [HotKeyIt/ahkdll-v2](https://github.com/hotkeyit/ahkdll/tree/alpha).

## Changes from ahkdll

- The object structure is the same as the AHK_L version.
- Remove standard library in autohotkey.exe resource, but libraries can still be loaded from resources.
- Remove the same hot string that defined in multiple threads firing at the same time.
- Remove `NULL`, `CriticalObject`, `Struct`, `sizeof`, `ObjLoad` and `ObjDump`
- Remove `A_ZipCompressionLevel`, specify this value in the last parameter of `ZipCreateBuffer` or `ZipCreateFile` or `ZipRawMemory`.
- `DynaCall` object has `Param[index]` property, used to retrieve and set default parameters, the index is the same as the position of the argument when it is called.
- `CryptAES` and Zip functions, parameter `Size` is not needed when previous parameter is an `Object` with `Ptr` and `Size` properties.
- `ahkExec(LPTSTR script, DWORD aThreadID = 0)` Inherits the current scope variable when aThreadID is `0`
- `Thread("Terminate", all := false)` Terminate one or all threads after the end of the current thread, and the timer is only terminating the current
- `FileRead` and `FileOpen` will detect `UTF8-RAW` when no code page is specified.
- Dynamic Library separator is `|` instead `:`, eg. `#include <urldownloadtovar|'http://www.website.com/script.ahk'>`, `urldownloadtovar.ahk` is a library in lib folder.
- Change `OnMessage(..., hwnd)` to `Gui.Prototype.OnMessage(Msg, Callback [, AddRemove])`, is similar to `Gui.Prototype.OnEvent(...)`
- Object literals support quoted property name, `v := {'key':val}`
- COM interface in dll module rather than in exe module.
- The com component has changed for the dll version, some methods are incorporated into [IAutoHotkeyLib](README-LIB.md#lib-api).
- Callback functions created through CallbackCreate that are called in other threads will be synchronized to the ahk thread through messages, which may produce unexpected results.
- Added `IAhkApi` export class for developing AHK bindings for third-party libraries, the header file is [ahkapi.h](https://github.com/thqby/AutoHotkey_H/blob/alpha/source/ahkapi.h).
- Added `__thiscall` calling conventions supported by `DllCall` and `DynaCall`, eg. `DllCall(func, params, 'thiscall')`, `DynaCall(func, 'ret_type=@params_type')`.
- Added `Decimal` class, supports arbitrary precision decimal operations, `MsgBox Decimal('0.1') + Decimal('0.2') = '0.3'`, [lib_mpir](https://github.com/thqby/AutoHotkey_H/releases/tag/lib_mpir) is required to compile, [the source code of mpir](https://github.com/wbhart/mpir)
- Added `Array.Prototype.Filter(callback: (value [, index]) => Boolean) => Array`
- Added `Array.Prototype.FindIndex(callback: (value [, index]) => Boolean, start_index := 1) => Integer`, if `start_index` less than 0 then reverse lookup
- Added `Array.Prototype.IndexOf(val_to_find, start_index := 1) => Integer`
- Added `Array.Prototype.Join(separator := ',') => String`
- Added `Array.Prototype.Map(callback: (value [, index]) => Any) => Array`
- Added `Array.Prototype.Sort(callback?: (a, b) => Integer) => $this`, sort in place and return and the default is random sort
- Added `#InitExec expression`, execute the expressions in load order before the script begins, is similar to the static initializer for v1
- Added `GuiControl.Prototype.OnMessage(Msg, Callback [, AddRemove])`, and `Gui.Prototype.OnMessage(Msg, Callback [, AddRemove])`, the parameter of the callback has changed, `Callback(GuiObj, wParam, lParam, Msg)`, `A_EventInfo` is the message posted time.
- Added `Object.Prototype.__Item[Prop]`, it's same as `Obj.%Prop%`
- Added `Object.Prototype.Get(Prop [, Default])`
- Support the use of `u8str`(utf8 string) in DllCall and `us` in DynaCall

## Classes List
```typescript
class Decimal {
  /**
   * Sets the computation precision and tostring() precision
   * @param prec Significant digits, greater than zero only affects division
   * @param outputprec tostring() If it is greater than 0, it is reserved to n decimal places. If less than 0, retain n significant digits
   * @return Returns the old prec value
   */
  static SetPrecision(prec := 20, outputprec := 0) => Integer

  // Converts integer, float, and numeric strings to Decimal object
  // Add, subtract, multiply and divide as with numbers
  static Call(val?) => Decimal

  ToString() => String

  // Convert to ahk value, integers outside the __int64 range are converted to double
  // Number(decimal_obj) => Integer | Float
}

class JSON {
  // JSON.stringify([JSON.null,JSON.true,JSON.false]) == '[null,true,false]'
  // !JSON.null == true
  // !JSON.false == true
  // JSON.false != false
  static null => ComValue
  static true => ComValue
  static false => ComValue

  /**
   * @param keep_type If true, convert true/false/null to JSON.true / JSON.false / JSON.null, otherwise 1 / 0 / ''
   * @param as_map If true, convert `{}` to Map, otherwise Object
   * supports [JSON5](https://spec.json5.org/) format.
   */
  static parse(text, keep_type := true, as_map := true) => Map | Array

  /**
   * the object include map,array,object and custom objects with `__enum` meta function
   * @param options {Integer|String|Object} The number of Spaces or string used for indentation
   * @param options.indent The number of spaces or string used for indentation
   * @param options.depth Expands the specified number of levels
   */
  static stringify(obj, options := 0) => String
}

class Worker {
  /**
   * Enumerates ahk threads.
   * @return {([&threadid,] &workerobj)=>void} An enumerator which will return items contained threadid and workerobj.
   */
  static __Enum(NumberOfVars?) => Enumerator

  /**
   * Creates a real AutoHotkey thread or associates an existing AutoHotkey thread in the current process and returns an object that communicates with it.
   * @param ScriptOrThreadID When ScriptOrThreadID is a script, create an AutoHotkey thread;
   * When ScriptOrThreadID is a threadid of created thread, it is associated with it;
   * When ScriptOrThreadID = 0, associate the main thread.
   */
  __New(ScriptOrThreadID := A_ThreadID, Cmd := '', Title := 'AutoHotkey') => Worker

  /**
   * Gets/sets the thread global variable. Objects of other threads will be converted to thread-safe Com object access and will not be accessible after the thread exits.
   * @param VarName Global variable name.
   */
  __Item[VarName] {
    get => Any
    set => void
  }

  /**
   * Call thread functions asynchronously. When the return value of another thread is an object, it is converted to a thread-safe Com object.
   * @param VarName The name of a global variable, or an object when it is associated with the current thread.
   * @param Params Parameters needed when called. The object type is converted to thread-safe Com object when passed to another thread.
   */
  AsyncCall(VarName, Params*) => Worker.Promise

  /**
   * Terminate the thread asynchronously.
   */
  ExitApp() => void

  /**
   * Pauses/Unpauses the script's current thread.
   */
  Pause(NewState) => Integer

  /**
   * Thread is ready.
   */
  Ready => Integer

  /**
   * Reload the thread asynchronously.
   */
  Reload() => void

  /**
   * Returns the thread ID.
   */
  ThreadID => Integer

  /**
   * Wait for the thread to exit, return 0 for timeout, or 1 otherwise.
   * @param Timeout The number of milliseconds, waitting until the thread exits when Timeout is 0.
   */
  Wait(Timeout := 0) => Integer

  class Promise extends Any {
    /**
     * Execute the callback after the asynchronous call completes.
     */
    Then(Callback) => Worker.Promise

    /**
     * An asynchronous call throws an exception and executes the callback.
     */
    Catch(Callback) => Worker.Promise
  }
}
```

## Functions List
```typescript
Alias(VarOrName [, VarOrPointer]) => void

Cast(DataType, Value, NewDataType) => Number

ComObjDll(hModule, CLSID [, IID]) => ComObject
 
CryptAES(AddOrBuf [, Size], password, EncryptOrDecrypt := true, Algorithm := 256) => Buffer

DynaCall(DllFunc, ParameterDefinition, Params*) => Number | String

GetVar(VarName, ResolveAlias := true) => Number

MemoryCallEntryPoint(hModule [, cmdLine]) => Number

MemoryFindResource(hModule, Name, Type [, Language]) => Number

MemoryFreeLibrary(hModule) => String

MemoryGetProcAddress(hModule, FuncName) => Number

MemoryLoadLibrary(PathToDll) => Number

MemoryLoadResource(hModule, hResource) => Number

MemoryLoadString(hModule, Id [, Language]) => String

MemorySizeOfResource(hModule, hReslnfo) => Number

ObjDump(obj [, compress, password]) => Buffer

ObjLoad(AddOrPath [, password]) => Array | Map | Object

ResourceLoadLibrary(ResName) => Number

Swap(Var1, Var2) => void

UArray(Values*) => Array

UMap([Key1, Value1, ...]) => Map

UObject([Key1, Value1, ...]) => Object

UnZip(AddOrBufOrFile [, Size], DestinationFolder, FileToExtract?, DestinationFileName?, Password?, CodePage := 0) => void

UnZipBuffer(AddOrBufOrFile [, Size], FileToExtract, Password?, CodePage := 0) => Buffer

UnZipRawMemory(AddOrBuf [, Size], Password?) => Buffer

ZipAddBuffer(ZipHandle, AddOrBuf [, Size], FileName?) => void

ZipAddFile(ZipHandle, FileName [, ZipFileName]) => void

ZipAddFolder(ZipHandle, ZipFoldName) => void

ZipCloseBuffer(ZipHandle) => Buffer

ZipCloseFile(ZipHandle) => void

ZipCreateBuffer(MaxSize, Password?, CompressionLevel := 5) => Number

ZipCreateFile(FileName, Password?, CompressionLevel := 5) => Number

ZipInfo(AddOrBufOrFile [, Size], CodePage := 0) => Array

ZipOptions(ZipHandle, Options) => void

ZipRawMemory(AddOrBuf [, Size], Password?, CompressionLevel := 5) => Buffer
```

## AutoHotkey.dll Module
### COM Interfaces
ProgID: `AutoHotkey2.Script`  
CLSID : `{934B0E6A-9B50-4248-884B-BE5A9BC66B39}`
The methods and properties exposed by the Lib object are defined in [ahklib.idl](source/ahklib.idl), in the  `IAutoHotkeyLib` interface.


### Export Functions
```cpp
PHOOK_ENTRY MinHookEnable(LPVOID pTarget, LPVOID pDetour, LPVOID *ppOriginal);

BOOL MinHookDisable(PHOOK_ENTRY pHook);

DWORD NewThread(LPCTSTR aScript, LPCTSTR aCmdLine = _T(""), LPCTSTR aTitle = _T("AutoHotkey"));

int ahkPause(int aNewState, DWORD aThreadID = 0);

UINT_PTR ahkFindLabel(LPTSTR aLabelName, DWORD aThreadID = 0);

LPTSTR ahkGetVar(LPTSTR name, int getVar = 0, DWORD aThreadID = 0);

int ahkAssign(LPTSTR name, LPTSTR value, DWORD aThreadID = 0);

UINT_PTR ahkExecuteLine(UINT_PTR line, int aMode, int wait, DWORD aThreadID = 0);

int ahkLabel(LPTSTR aLabelName, int nowait = 0, DWORD aThreadID = 0);

UINT_PTR ahkFindFunc(LPTSTR funcname, DWORD aThreadID = 0);

LPTSTR ahkFunction(LPTSTR func, LPTSTR param1 = NULL, LPTSTR param2 = NULL, LPTSTR param3 = NULL, LPTSTR param4 = NULL,     LPTSTR param5 = NULL, LPTSTR param6 = NULL, LPTSTR param7 = NULL, LPTSTR param8 = NULL, LPTSTR param9 = NULL, LPTSTR param10 = NULL, DWORD aThreadID = 0);

int ahkPostFunction(LPTSTR func, LPTSTR param1 = NULL, LPTSTR param2 = NULL, LPTSTR param3 = NULL, LPTSTR param4 = NULL, LPTSTR param5 = NULL, LPTSTR param6 = NULL, LPTSTR param7 = NULL, LPTSTR param8 = NULL, LPTSTR param9 = NULL, LPTSTR param10 = NULL, DWORD aThreadID = 0);

int ahkReady(DWORD aThreadID = 0);

UINT_PTR addScript(LPTSTR script, int waitexecute = 0, DWORD aThreadID = 0);

int ahkExec(LPTSTR script, DWORD aThreadID = 0);
```

## How to Compile

AutoHotkey is developed with [Microsoft Visual Studio Community 2022](https://www.visualstudio.com/products/visual-studio-community-vs), which is a free download from Microsoft.

  - Get the source code.
  - Open AutoHotkeyx.sln in Visual Studio.
  - Select the appropriate Build and Platform.
  - Build.

The project is configured in a way that allows building with Visual Studio 2012 or later, but only the 2022 toolset is regularly tested. Some newer C++ language features are used and therefore a later version of the compiler might be required.


## Developing in VS Code ##

AutoHotkey v2 can also be built and debugged in VS Code.

Requirements:
  - [C/C++ for Visual Studio Code](https://marketplace.visualstudio.com/items?itemName=ms-vscode.cpptools). VS Code might prompt you to install this if you open a .cpp file.
  - [Build Tools for Visual Studio 2022](https://aka.ms/vs/17/release/vs_BuildTools.exe) with the "Desktop development with C++" workload, or similar (some older or newer versions and different products should work).



## Build Configurations ##

AutoHotkeyx.vcxproj contains several combinations of build configurations.  The main configurations are:

  - **Debug**: AutoHotkey.exe in debug mode.
  - **Release**: AutoHotkey.exe for general use.
  - **Self-contained**: AutoHotkeySC.bin, used for compiled scripts.

Secondary configurations are:

  - **(mbcs)**: ANSI (multi-byte character set). Configurations without this suffix are Unicode.
  - **.dll**: Builds an experimental dll for use hosting the interpreter, such as to enable the use of v1 libraries in a v2 script. See [README-LIB.md](README-LIB.md).


## Platforms ##

AutoHotkeyx.vcxproj includes the following Platforms:

  - **Win32**: for Windows 32-bit.
  - **x64**: for Windows x64.

AutoHotkey supports Windows XP with or without service packs and Windows 2000 via an asm patch (win2kcompat.asm).  Support may be removed if maintaining it becomes non-trivial.  Older versions are not supported.
