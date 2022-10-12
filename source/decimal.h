#pragma once
#ifdef ENABLE_DECIMAL

#include "lib_mpir/mpir.h"

class Decimal : public ObjectBase
{
	mpz_t z;
	__int64 e;

	static const mpir_si sCarryPrec = -4;
	thread_local static mpir_si sPrec;
	thread_local static mpir_si sOutputPrec;

	void copy_to(Decimal *to) { mpz_set(to->z, z); to->e = e; }
	void assign(__int64 value) { mpz_set_sx(z, value); }
	void assign(double value);
	bool assign(LPCTSTR str);
	void assign(const char *str, size_t len, int base = 10);
	void carry(bool ignore_integer = true, bool fix = false);
	bool is_integer();
	LPTSTR to_string();

	static void mul_10exp(Decimal *v, Decimal *a, mpir_ui e);
	static void add_or_sub(Decimal *v, Decimal *a, Decimal *b, int add = 1);
	static void mul(Decimal *v, Decimal *a, Decimal *b);
	static int div(Decimal *v, Decimal *a, Decimal *b, bool intdiv = false);

public:
	enum MemberID
	{
		M_ToString,
	};

	Decimal() { mpz_init(z); e = 0; }
	~Decimal() { mpz_clear(z); }

	static Decimal *Create(ExprTokenType *aToken = nullptr);
	static void Create(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount);
	static void SetPrecision(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount);
	int Eval(ExprTokenType &op, ExprTokenType *right = nullptr);
	void Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	BOOL ToBOOL() { return z->_mp_size != 0; }
	static Decimal *ToDecimal(IObject *obj) { return *(void **)obj == sVTable ? static_cast<Decimal *>(obj) : nullptr; }
	static Decimal *ToDecimal(ExprTokenType &aToken);
	ResultType ToToken(ExprTokenType &aToken);

	IObject_Type_Impl("Decimal");
	::Object *Base() { return sPrototype; }
	static void *getVTable() { Decimal t; return *(void **)&t; }
	static void *sVTable;
	static ObjectMember sMembers[];
	thread_local static Object *sPrototype;
};


#endif // ENABLE_DECIMAL