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
- Remove `#UseStdLib`, and automatically load libraries from resources.
- Remove the same hot string that defined in multiple threads firing at the same time.
- Remove `NULL`, `CriticalObject`, `Struct`, `sizeof`, `ObjLoad` and `ObjDump`
- Remove `A_ZipCompressionLevel`, specify this value in the last parameter of `ZipCreateBuffer` or `ZipCreateFile` or `ZipRawMemory`.
- `DynaCall` object has `Param[index]` property, used to retrieve and set default parameters, the index is the same as the position of the argument when it is called.
- `CryptAES` and Zip functions, parameter `Size` is not needed when previous parameter is an `Object` with `Ptr` and `Size` properties.
- `ahkExec(LPTSTR script, DWORD aThreadID = 0)` Inherits the current scope variable when aThreadID is `0`
- `Thread("Terminate", all := false)` Terminate the last or all threads, and the timer is only terminating the current
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
- Added `GuiControl.Prototype.OnMessage(Msg, Callback [, AddRemove])`, and `Gui.Prototype.OnMessage(Msg, Callback [, AddRemove])`, the parameter of the callback has changed, `Callback(GuiObj, wParam, lParam, Msg)`, `A_EventInfo` is the message posted time.
- Added `Object.Prototype.__Item[Prop]`, it's same as `Obj.%Prop%`
- Added `Object.Prototype.Get(Prop [, Default])`
- Support the use of `u8str`(utf8 string) in DllCall and `us` in DynaCall

## [Classes/Functions List](https://github.com/thqby/vscode-autohotkey2-lsp/blob/main/syntaxes/ahk2_h.d.ahk)

## AutoHotkey.dll Module
### COM Interfaces
ProgID: `AutoHotkey2.Script`  
CLSID : `{934B0E6A-9B50-4248-884B-BE5A9BC66B39}`
The methods and properties exposed by the Lib object are defined in [ahklib.idl](source/ahklib.idl), in the  `IAutoHotkeyLib` interface.

### [Export Functions](https://github.com/thqby/AutoHotkey_H/blob/alpha-h/source/ahkapi.h#L457-L469)

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
