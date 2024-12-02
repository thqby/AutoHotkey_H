#pragma once


struct ScriptImport
{
	LPTSTR names = nullptr;
	ScriptModule *mod = nullptr;
	ScriptImport *next = nullptr;
	LineNumberType line_number = 0;
	FileIndexType file_index = 0;

	ScriptImport() {}
	ScriptImport(ScriptModule *aMod) : mod(aMod), names(_T("*")) {}

	void *operator new(size_t aBytes) {return SimpleHeap::Alloc(aBytes);}
	void *operator new[](size_t aBytes) {return SimpleHeap::Alloc(aBytes);}
	void operator delete(void *aPtr) {}
	void operator delete[](void *aPtr) {}
};


class ScriptModule : public ObjectBase
{
public:
	LPCTSTR mName = nullptr;
	Line *mFirstLine = nullptr;
	Label *mFirstLabel = nullptr;
	ScriptImport *mImports = nullptr;
	ScriptModule *mPrev = nullptr;
	VarList mVars;
	Var *mSelf = nullptr;
	UnresolvedBaseClass *mUnresolvedBaseClass = nullptr;
	FileIndexType *mFiles = nullptr, mFilesCount = 0, mFilesCountMax = 0;
	bool mExecuted = false;
	bool mIsBuiltinModule = false;

	// #Warn settings
	WarnMode Warn_LocalSameAsGlobal = WARNMODE_OFF;
	WarnMode Warn_Unreachable = WARNMODE_ON;
	WarnMode Warn_VarUnset = WARNMODE_ON;

	bool HasFileIndex(FileIndexType aFile);
	ResultType AddFileIndex(FileIndexType aFile);

	IObject *FindGlobalObject(LPCTSTR aName);

	ScriptModule() {}
	ScriptModule(LPCTSTR aName) : mName(aName) {}

	void *operator new(size_t aBytes) {return SimpleHeap::Alloc(aBytes);}
	void *operator new[](size_t aBytes) {return SimpleHeap::Alloc(aBytes);}
	void operator delete(void *aPtr) {}
	void operator delete[](void *aPtr) {}

	IObject_Type_Impl("Module");
	ResultType Invoke(IObject_Invoke_PARAMS_DECL) override;
	Object *Base() override { return sPrototype; }
	static Object *sPrototype;
};

typedef ScriptItemList<ScriptModule, 16> ScriptModuleList;
