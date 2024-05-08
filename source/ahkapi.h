#ifndef AHK2_TYPES_H
#define AHK2_TYPES_H
#include <OAIdl.h>

#define CONFIG_DEBUGGER

#ifndef _T
#ifdef _UNICODE
#define _T(x)	L ## x
#else
#define _T(x)	x
#endif
#endif

#define IT_GET				0
#define IT_SET				1
#define IT_CALL				2
#define IT_BITMASK			3

#define INVOKE_TYPE			(aFlags & IT_BITMASK)
#define IS_INVOKE_SET		(aFlags & IT_SET)
#define IS_INVOKE_GET		(INVOKE_TYPE == IT_GET)
#define IS_INVOKE_CALL		(aFlags & IT_CALL)

#define INVOKE_NOT_HANDLED	CONDITION_FALSE

#define MAX_NUMBER_LENGTH 255
#define MAX_NUMBER_SIZE (MAX_NUMBER_LENGTH + 1)
#define MAX_INTEGER_LENGTH 20
#define MAX_INTEGER_SIZE (MAX_INTEGER_LENGTH + 1)

#define _f_callee_id			(aResultToken.func->mFID)


enum ResultType {
	FAIL = 0, OK, WARN = OK, CRITICAL_ERROR
	, CONDITION_TRUE, CONDITION_FALSE
	, LOOP_BREAK, LOOP_CONTINUE
	, EARLY_RETURN, EARLY_EXIT
	, FAIL_OR_OK
};

enum SymbolType
{
	PURE_NOT_NUMERIC
	, PURE_INTEGER, PURE_FLOAT
	, SYM_STRING = PURE_NOT_NUMERIC, SYM_INTEGER = PURE_INTEGER, SYM_FLOAT = PURE_FLOAT
#define IS_NUMERIC(symbol) ((symbol) == SYM_INTEGER || (symbol) == SYM_FLOAT)
	, SYM_MISSING
	, SYM_VAR
	, SYM_OBJECT
	, SYM_DYNAMIC
	, SYM_TYPED_FIELD = 76
};


//const IID IID_IObjectComCompatible = { 0x619f7e25, 0x6d89, 0x4eb4, 0xb2, 0xfb, 0x18, 0xe7, 0xc7, 0x3c, 0xe, 0xa6 };

struct ExprTokenType;
struct ResultToken;
class Object;

struct DECLSPEC_NOVTABLE IObject : public IDispatch
{
#define IObject_Invoke_PARAMS_DECL ResultToken &aResultToken, int aFlags, LPTSTR aName, ExprTokenType &aThisToken, ExprTokenType *aParam[], int aParamCount
#define IObject_Invoke_PARAMS aResultToken, aFlags, aName, aThisToken, aParam, aParamCount
	virtual ResultType Invoke(IObject_Invoke_PARAMS_DECL) = 0;
	virtual LPTSTR Type() = 0;
#define IObject_Type_Impl LPTSTR Type() { return _T(CLASSNAME); }
	virtual Object* Base() = 0;
	virtual bool IsOfType(Object* aPrototype) = 0;

#ifdef CONFIG_DEBUGGER
#define IObject_DebugWriteProperty_Def void DebugWriteProperty(void *aDebugger, int aPage, int aPageSize, int aMaxDepth)
	virtual void DebugWriteProperty(void* aDebugger, int aPage, int aPageSize, int aMaxDepth) = 0;
#else
#define IObject_DebugWriteProperty_Def
#endif
};

#define MAXP_VARIADIC 255

template <typename T, typename index_t = ::size_t>
class FlatVector
{
public:
	struct Data
	{
		index_t size;
		index_t length;
	};
	Data* data = &Empty;
	struct OneT : public Data { char zero_buf[sizeof(T)]; };
	static OneT Empty;

	index_t& Length() { return data->length; }
	index_t Capacity() { return data->size; }
	T* Value() { return (T*)(data + 1); }
	operator T* () { return Value(); }
};

template <typename T, typename index_t>
typename FlatVector<T, index_t>::OneT FlatVector<T, index_t>::Empty;


class Property;

class Var;
class Func;
class BuiltInFunc;

struct ExprTokenType
{
	union
	{
		__int64 value_int64;
		double value_double;
		struct
		{
			union
			{
				IObject* object;
				Var* var;
				LPTSTR marker;
			};
			union
			{
				size_t marker_length;
				int var_usage;
			};
		};
	};
	SymbolType symbol;
	ExprTokenType() : value_int64(0)
#ifdef _WIN64
		, marker_length(0)
#endif
	{}
	ExprTokenType(LPTSTR str) { SetValue(str); }
	ExprTokenType(IObject* obj) { SetValue(obj); }
	ExprTokenType(int val) { SetValue((__int64)val); }
	ExprTokenType(__int64 val) { SetValue(val); }
	ExprTokenType(double val) { SetValue(val); }
	void SetValue(LPTSTR str, size_t len = -1) {
		marker = str, marker_length = len, symbol = SYM_STRING;
	}
	void SetValue(int val) { SetValue((__int64)val); }
	void SetValue(__int64 val) {
		value_int64 = val, symbol = SYM_INTEGER;
#ifdef _WIN64
		marker_length = 0;
#endif
	}
	void SetValue(double val) {
		value_double = val, symbol = SYM_FLOAT;
#ifdef _WIN64
		marker_length = 0;
#endif
	}
	void SetValue(IObject* obj) {
		object = obj, symbol = SYM_OBJECT;
#ifdef _WIN64
		marker_length = 0;
#endif
	}
};

struct ResultToken : public ExprTokenType
{
	LPTSTR buf;
	LPTSTR mem_to_free;
#ifdef ENABLE_HALF_BAKED_NAMED_PARAMS
	IObject* named_params;
#endif
	BuiltInFunc* func;

	bool Exited()
	{
		return result == FAIL || result == EARLY_EXIT;
	}

	void AcceptMem(LPTSTR aNewMem, size_t aLength)
	{
		symbol = SYM_STRING;
		marker = mem_to_free = aNewMem;
		marker_length = aLength;
	}
	ResultType result;
};

typedef UCHAR VarTypeType;
typedef UCHAR AllocMethodType;
typedef UCHAR VarAttribType;
typedef UINT_PTR VarSizeType;

#pragma warning(push)
#ifdef _WIN64
#pragma pack(push, 8)
#else
#pragma pack(push, 4)
#endif
class Var
{
public:
	union
	{
		__int64 mContentsInt64;
		double mContentsDouble;
		IObject* mObject;
	};
	union
	{
		LPTSTR mCharContents;
		char* mByteContents;
	};
	union
	{
		Var* mAliasFor = nullptr;
		VarSizeType mByteLength;
	};
	VarSizeType mByteCapacity = 0;
	AllocMethodType mHowAllocated = 0;
	VarAttribType mAttrib;
	UCHAR mScope;
	VarTypeType mType;
	TCHAR* mName;
};
#pragma pack(pop)
#pragma warning(pop)

class ObjectBase : public IObject
{
public:
	ULONG mRefCount;
#ifdef _WIN64
	UINT mFlags;
#endif
	virtual bool Delete() = 0;
public:
	virtual ~ObjectBase() = 0;
};

class VarRef : public ObjectBase, public Var {};

class Object : public ObjectBase
{
public:
#ifndef _WIN64
	UINT mFlags;
#endif

	typedef LPTSTR name_t;
	typedef FlatVector<TCHAR> String;

	typedef UINT index_t;
	struct Variant
	{
		union {
			__int64 n_int64;
			double n_double;
			IObject* object;
			String string;
			Property* prop;
		};
		SymbolType symbol;
		TCHAR key_c;
	};

	struct FieldType : Variant
	{
		name_t name;
	};

	enum EnumeratorType
	{
		Enum_Properties,
		Enum_Methods
	};

	enum Flags : decltype(mFlags)
	{
		UnsortedFlag = 0x80000000,
		ClassPrototype = 0x01,
		NativeClassPrototype = 0x02,
		DataIsSetFlag = 0x04,
		DataIsAllocatedFlag = 0x08,
		DataIsStructInfo = 0x10,
		StructInfoLocked = 0x20,
		LastObjectFlag = 0x20
	};

	Object* mBase = nullptr;
	FlatVector<FieldType, index_t> mFields;
	void *mData = nullptr;
	Object **mNested = nullptr;
};
class Array : public Object
{
public:
	Variant* mItem = nullptr;
	index_t mLength = 0, mCapacity = 0;

	enum : index_t
	{
		BadIndex = UINT_MAX,
		MaxIndex = INT_MAX
	};
};

typedef __int64 IntKeyType;

class Map : public Object
{
public:
	union Key
	{
		LPTSTR s;
		IntKeyType i;
		IObject* p;
	};
	struct Pair : Variant
	{
		Key key;
	};

	enum MapOption : decltype(mFlags)
	{
		MapCaseless = LastObjectFlag << 1,
			MapUseLocale = MapCaseless << 1
	};

	Pair* mItem = nullptr;
	index_t mCount = 0, mCapacity = 0;

	static const index_t mKeyOffsetInt = 0;
	index_t mKeyOffsetObject = 0, mKeyOffsetString = 0;
};

class BufferObject : public Object
{
public:
	void* mData;
	size_t mSize;
};

class ComObject : public ObjectBase
{
public:
	union
	{
		IDispatch* mDispatch;
		IUnknown* mUnknown;
		SAFEARRAY* mArray;
		void* mValPtr;
		__int64 mVal64;
	};
	void* mEventSink;
	VARTYPE mVarType;
	enum { F_OWNVALUE = 1 };
	USHORT mFlags;
};

class DECLSPEC_NOVTABLE Func : public Object
{
public:
	LPCTSTR mName;
	int mParamCount = 0;
	int mMinParams = 0;
	bool mIsVariadic = false;

	virtual bool IsBuiltIn() = 0;
	virtual bool ArgIsOutputVar(int aArg) = 0;
	virtual bool ArgIsOptional(int aArg) = 0;
	virtual bool Call(ResultToken& aResultToken, ExprTokenType* aParam[], int aParamCount) = 0;
};

class DECLSPEC_NOVTABLE EnumBase : public Func
{
public:
	virtual ResultType Next(Var*, Var*) = 0;
};

class ComArrayEnum : public EnumBase
{
	ComObject* mArrayObject;
	void* mData;
	long mLBound, mUBound;
	UINT mElemSize;
	VARTYPE mType;
	bool mIndexMode;
	long mOffset;
};

class DECLSPEC_NOVTABLE NativeFunc : public Func
{
};

#define BIF_DECL_PARAMS ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount
#define BIF_DECL(name) void name(BIF_DECL_PARAMS)

typedef BIF_DECL((*BuiltInFunctionType));
class BuiltInFunc : public NativeFunc
{
public:
	BuiltInFunctionType mBIF;
	union {
		int mFID;
		void* mData;
	};
	UCHAR *mOutputVars;
};

#define MAX_FUNC_OUTPUT_VAR 7
struct FuncEntry
{
	LPCTSTR mName;
	BuiltInFunctionType mBIF;
	UCHAR mMinParams, mMaxParams;
	UCHAR mID;
	UCHAR mOutputVars[MAX_FUNC_OUTPUT_VAR];
};

#define MAX_FUNCTION_PARAMS UCHAR_MAX
#define NA MAX_FUNCTION_PARAMS
#define BIFn(name, minp, maxp, bif, ...) {_T(#name), bif, minp, maxp, FID_##name, __VA_ARGS__}
#define BIFi(name, minp, maxp, bif, id, ...) {_T(#name), bif, minp, maxp, id, __VA_ARGS__}
#define BIF1(name, minp, maxp, ...) {_T(#name), BIF_##name, minp, maxp, 0, __VA_ARGS__}

typedef void (IObject::* ObjectMethod)(ResultToken& aResultToken, int aID, int aFlags, ExprTokenType* aParam[], int aParamCount);
class BuiltInMethod : public NativeFunc
{
public:
	ObjectMethod mBIM;
	Object* mClass;
	UCHAR mMID;
	UCHAR mMIT;
};

struct ObjectMember
{
	LPTSTR name;
	ObjectMethod method;
	UCHAR id, invokeType, minParams, maxParams;
};

#define Object_Member(name, impl, id, invokeType, ...) \
	{ _T(#name), static_cast<ObjectMethod>(&impl), id, invokeType, __VA_ARGS__ }
#define Object_Method_(name, minP, maxP, impl, id) Object_Member(name, impl,   id,       IT_CALL, minP, maxP)
#define Object_Method(name, minP, maxP)            Object_Member(name, Invoke, M_##name, IT_CALL, minP, maxP)
#define Object_Method1(name, minP, maxP)           Object_Member(name, name,   0,        IT_CALL, minP, maxP)
#define Object_Property_get(name, ...)             Object_Member(name, Invoke, P_##name, IT_GET, __VA_ARGS__)
#define Object_Property_get_set(name, ...)         Object_Member(name, Invoke, P_##name, IT_SET, __VA_ARGS__)


#define EXPORT

EXPORT DWORD NewThread(LPTSTR aScript, LPTSTR aCmdLine = NULL, LPTSTR aTitle = NULL);
EXPORT UINT_PTR addScript(LPTSTR script, int waitexecute = 0, DWORD aThreadID = 0);
EXPORT UINT_PTR ahkExecuteLine(UINT_PTR line, int aMode, int wait, DWORD aThreadID = 0);
EXPORT UINT_PTR ahkFindFunc(LPTSTR aFuncName, DWORD aThreadID = 0);
EXPORT UINT_PTR ahkFindLabel(LPTSTR aLabelName, DWORD aThreadID = 0);
EXPORT LPTSTR ahkFunction(LPTSTR func, LPTSTR param1 = NULL, LPTSTR param2 = NULL, LPTSTR param3 = NULL, LPTSTR param4 = NULL, LPTSTR param5 = NULL, LPTSTR param6 = NULL, LPTSTR param7 = NULL, LPTSTR param8 = NULL, LPTSTR param9 = NULL, LPTSTR param10 = NULL, DWORD aThreadID = 0);
EXPORT LPTSTR ahkGetVar(LPTSTR name, int getVar = 0, DWORD aThreadID = 0);
EXPORT int ahkAssign(LPTSTR name, LPTSTR value, DWORD aThreadID = 0);
EXPORT int ahkExec(LPTSTR script, DWORD aThreadID = 0);
EXPORT int ahkLabel(LPTSTR aLabelName, int nowait = 0, DWORD aThreadID = 0);
EXPORT int ahkPause(LPTSTR aChangeTo, DWORD aThreadID = 0);
EXPORT int ahkPostFunction(LPTSTR func, LPTSTR param1 = NULL, LPTSTR param2 = NULL, LPTSTR param3 = NULL, LPTSTR param4 = NULL, LPTSTR param5 = NULL, LPTSTR param6 = NULL, LPTSTR param7 = NULL, LPTSTR param8 = NULL, LPTSTR param9 = NULL, LPTSTR param10 = NULL, DWORD aThreadID = 0);
EXPORT int ahkReady(DWORD aThreadID = 0);

class IAhkApi : public IUnknown
{
public:
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
		MAXINDEX
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
	virtual void STDMETHODCALLTYPE VarFree(Var* aVar, bool aExcludeAliasesAndRequireInit = false);
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
};

EXPORT IAhkApi* ahkGetApi(void* options = nullptr);
#endif