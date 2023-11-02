#ifndef exports_h
#define exports_h

#define EXPORT(_rettype_) extern "C" _rettype_ __stdcall

EXPORT(DWORD) NewThread(LPTSTR aScript, LPTSTR aCmdLine = NULL, LPTSTR aTitle = NULL);
EXPORT(UINT_PTR) addScript(LPTSTR script, int waitexecute = 0, DWORD aThreadID = 0);
EXPORT(UINT_PTR) ahkExecuteLine(UINT_PTR line, int aMode, int wait, DWORD aThreadID = 0);
EXPORT(UINT_PTR) ahkFindFunc(LPTSTR aFuncName, DWORD aThreadID = 0);
EXPORT(UINT_PTR) ahkFindLabel(LPTSTR aLabelName, DWORD aThreadID = 0);
EXPORT(LPTSTR) ahkFunction(LPTSTR func, LPTSTR param1 = NULL, LPTSTR param2 = NULL, LPTSTR param3 = NULL, LPTSTR param4 = NULL, LPTSTR param5 = NULL, LPTSTR param6 = NULL, LPTSTR param7 = NULL, LPTSTR param8 = NULL, LPTSTR param9 = NULL, LPTSTR param10 = NULL, DWORD aThreadID = 0);
EXPORT(LPTSTR) ahkGetVar(LPTSTR name, int getVar = 0, DWORD aThreadID = 0);
EXPORT(int) ahkAssign(LPTSTR name, LPTSTR value, DWORD aThreadID = 0);
EXPORT(int) ahkExec(LPTSTR script, DWORD aThreadID = 0);
EXPORT(int) ahkLabel(LPTSTR aLabelName, int nowait = 0, DWORD aThreadID = 0);
EXPORT(int) ahkPause(int aNewState, DWORD aThreadID = 0);
EXPORT(int) ahkPostFunction(LPTSTR func, LPTSTR param1 = NULL, LPTSTR param2 = NULL, LPTSTR param3 = NULL, LPTSTR param4 = NULL, LPTSTR param5 = NULL, LPTSTR param6 = NULL, LPTSTR param7 = NULL, LPTSTR param8 = NULL, LPTSTR param9 = NULL, LPTSTR param10 = NULL, DWORD aThreadID = 0);
EXPORT(int) ahkReady(DWORD aThreadID = 0);
DWORD NewThread(LPTSTR aScript, LPTSTR aCmdLine, LPTSTR aTitle, void(*callback)(void *param), void *param);

class IAhkApi : public IUnknown
{
public:
	static IAhkApi* Initialize();
	static void Finalize();
		
	enum class ObjectType {
		Object,
		Array,
		Buffer,
		File,
		Map,
		VarRef,
		ComValue,
		ComObjArray,
		ComObject,
		UObject,
		UArray,
		UMap,
		Worker,
		Module,
		MAXINDEX,
	};

	enum class ErrorType
	{
		Error, Index, Member, Property, Method, Memory, OS,
		Target, Timeout, Type, Unset, UnsetItem, Value, ZeroDivision
	};

	class Prototype : public Object
	{
	public:
		size_t mSize = 0;
		void (*OnDispose)() = nullptr;
		~Prototype();
		virtual Object* New(ExprTokenType* aParam[], int aParamCount);
	};

	// IUnknown
	ULONG STDMETHODCALLTYPE AddRef();
	ULONG STDMETHODCALLTYPE Release();
	STDMETHODIMP QueryInterface(REFIID riid, void** ppv);

	virtual void* STDMETHODCALLTYPE Malloc(size_t aSize);
	virtual void* STDMETHODCALLTYPE Realloc(void* aPtr, size_t aSize);
	virtual void STDMETHODCALLTYPE Free(void* aPtr);

	// Don't need free when return by `ExprTokenType` or call `TokenToString`
	virtual LPTSTR STDMETHODCALLTYPE TokenToString(ExprTokenType& aToken, TCHAR aBuf[] = nullptr, size_t* aLength = nullptr);
	virtual bool STDMETHODCALLTYPE TokenToNumber(ExprTokenType& aInput, ExprTokenType& aOutput);
	virtual bool STDMETHODCALLTYPE VarAssign(Var* aVar, ExprTokenType& aToken);
	virtual void STDMETHODCALLTYPE VarToToken(Var* aVar, ExprTokenType& aToken);
	virtual void STDMETHODCALLTYPE VarFree(Var* aVar, int aWhenToFree = VAR_ALWAYS_FREE | VAR_CLEAR_ALIASES);
	virtual bool STDMETHODCALLTYPE VariantAssign(Object::Variant& aVariant, ExprTokenType& aValue);
	virtual void STDMETHODCALLTYPE VariantToToken(Object::Variant& aVariant, ExprTokenType& aToken);
	virtual void STDMETHODCALLTYPE VariantToToken(VARIANT& aVariant, ResultToken& aToken, bool aRetainVar = true);
	virtual void STDMETHODCALLTYPE ResultTokenFree(ResultToken& aToken);
	virtual ResultType STDMETHODCALLTYPE Error(LPTSTR aErrorText, LPTSTR aExtraInfo = _T(""), ErrorType aType = ErrorType::Error);
	virtual ResultType STDMETHODCALLTYPE TypeError(LPTSTR aExpectedType, ExprTokenType& aToken);

	virtual void* STDMETHODCALLTYPE GetProcAddress(LPTSTR aDllFileFunc, HMODULE* hmodule_to_free = nullptr);
	virtual void* STDMETHODCALLTYPE GetProcAddressCrc32(HMODULE aModule, UINT aCRC32, UINT aInitial = 0);

	virtual bool STDMETHODCALLTYPE Script_GetVar(LPTSTR aVarName, ExprTokenType& aValue);
	virtual bool STDMETHODCALLTYPE Script_SetVar(LPTSTR aVarName, ExprTokenType& aValue);

	virtual Func* STDMETHODCALLTYPE Func_New(FuncEntry& aBIF);
	virtual Func* STDMETHODCALLTYPE Method_New(LPTSTR aFullName, ObjectMember& aMember, Object* aPrototype);
	// `aPrototype` does not need to be released, `obj->IsOfType(aPrototype)` is used to check the object type
	virtual Object* STDMETHODCALLTYPE Class_New(LPTSTR aClassName, size_t aClassSize, ObjectMember aMembers[], int aMemberCount, Prototype*& aPrototype, Object* aBase = nullptr);

	virtual IObject* STDMETHODCALLTYPE GetEnumerator(IObject* aObj, int aVarCount);
	virtual bool STDMETHODCALLTYPE CallEnumerator(IObject* aEnumerator, ExprTokenType* aParam[], int aParamCount);

	virtual IObject* STDMETHODCALLTYPE Object_New(ObjectType aType = ObjectType::Object, ExprTokenType* aParam[] = nullptr, int aParamCount = 0);
	virtual bool STDMETHODCALLTYPE Object_CallProp(ResultToken& aResultToken, Object* aObject, LPTSTR aName, ExprTokenType* aParam[], int aParamCount);
	virtual bool STDMETHODCALLTYPE Object_GetProp(ResultToken& aResultToken, Object* aObject, LPTSTR aName, bool aOwnProp = false, ExprTokenType* aParam[] = nullptr, int aParamCount = 0);
	virtual bool STDMETHODCALLTYPE Object_SetProp(Object* aObject, LPTSTR aName, ExprTokenType& aValue, bool aOwnProp = false, ExprTokenType* aParam[] = nullptr, int aParamCount = 0);
	virtual bool STDMETHODCALLTYPE Object_DeleteOwnProp(ResultToken& aResultToken, Object* aObject, LPTSTR aName);
	virtual void STDMETHODCALLTYPE Object_Clear(Object* aObject);

	virtual bool STDMETHODCALLTYPE Array_GetItem(Array* aArray, UINT aIndex, ExprTokenType& aValue);
	virtual bool STDMETHODCALLTYPE Array_SetItem(Array* aArray, UINT aIndex, ExprTokenType& aValue);
	virtual bool STDMETHODCALLTYPE Array_InsertItem(Array* aArray, ExprTokenType& aValue, UINT* aIndex = nullptr);
	virtual bool STDMETHODCALLTYPE Array_DeleteItem(ResultToken& aResultToken, Array* aArray, UINT aIndex);
	virtual bool STDMETHODCALLTYPE Array_RemoveItems(Array* aArray, UINT aIndex = 0, UINT aCount = -1);
	virtual Array* STDMETHODCALLTYPE Array_FromEnumerable(IObject* aEnumerable, UINT aIndex = 0);

	virtual bool STDMETHODCALLTYPE Buffer_Resize(BufferObject* aBuffer, size_t aSize);
	
	virtual bool STDMETHODCALLTYPE Map_GetItem(Map* aMap, ExprTokenType& aKey, ExprTokenType& aValue);
	virtual bool STDMETHODCALLTYPE Map_SetItem(Map* aMap, ExprTokenType& aKey, ExprTokenType& aValue);
	virtual bool STDMETHODCALLTYPE Map_DeleteItem(ResultToken& aResultToken, Map* aMap, ExprTokenType& aKey);
	virtual void STDMETHODCALLTYPE Map_Clear(Map* aMap);

	virtual Object* STDMETHODCALLTYPE JSON_Parse(LPTSTR aJSON);
	virtual LPTSTR STDMETHODCALLTYPE JSON_Stringify(IObject* aObject, LPTSTR aIndent = nullptr);

	virtual void STDMETHODCALLTYPE PumpMessages();

	virtual Func *STDMETHODCALLTYPE MdFunc_New(LPCTSTR aName, void* aFuncPtr, MdType* aSig, Object *aPrototype = nullptr);

private:
	thread_local static Object* sObject[(int)ObjectType::MAXINDEX];
	thread_local static int sInit;
	static IAhkApi instance;
	IAhkApi() {};
};

EXPORT(IAhkApi*) ahkGetApi(void* options = nullptr);
#endif