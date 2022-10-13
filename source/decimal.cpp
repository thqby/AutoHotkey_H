#include "stdafx.h" // pre-compiled headers#include "defines.h"
#ifdef ENABLE_DECIMAL
#include "defines.h"
#include "script_object.h"
#include "decimal.h"
#include "var.h"
#include "script.h"

#ifdef _WIN64
#pragma comment(lib, "source/lib_mpir/Win64/mpir.lib")
#else
#pragma comment(lib, "source/lib_mpir/Win32/mpir.lib")
#endif // _WIN64



struct bases
{
	int chars_per_limb;
	double chars_per_bit_exactly;
	mp_limb_t big_base;
	mp_limb_t big_base_inverted;
};
extern "C" const struct bases __gmpn_bases[];
extern "C" const unsigned char __gmp_digit_value_tab[];

void Decimal::assign(double value)
{
	const int kDecimalRepCapacity = double_conversion::DoubleToStringConverter::kBase10MaximalLength + 1;
	char decimal_rep[kDecimalRepCapacity];
	int decimal_rep_length;
	int decimal_point;
	bool sign;
	double_conversion::DoubleToStringConverter::DoubleToAscii(value,
		double_conversion::DoubleToStringConverter::DtoaMode::SHORTEST,
		0, decimal_rep, kDecimalRepCapacity,
		&sign, &decimal_rep_length, &decimal_point);
	if (decimal_rep_length >= 16) {
		auto len = 15;
		char rep[kDecimalRepCapacity];
		for (int i = 0; i < decimal_rep_length; i++)
			rep[i] = decimal_rep[i];
		auto p = &rep[len - 1];
		if (*p == '0')
			len--;
		else if (rep[len] > '4') {
			rep[len] = '0';
			for ((*p)++; p > rep;)
				if (*p > '9')
					*p = '0', (*--p)++;
				else break;
			if (*p > '9')
				rep[0] = '1', len = 1, decimal_point++;
		}
		while (len > 1 && rep[len - 1] == '0')
			len--;
		if (decimal_rep_length - len > 4)
			for (decimal_rep[decimal_rep_length = len--] = 0; len >= 0; len--)
				decimal_rep[len] = rep[len];
	}
	for (int i = 0; i < decimal_rep_length; i++)
		decimal_rep[i] -= '0';
	assign(decimal_rep, decimal_rep_length);
	if (sign)
		z->_mp_size = -z->_mp_size;
	e = decimal_point - decimal_rep_length;
}

void Decimal::assign(const char *str, size_t len, int base)
{
	auto sz = (((mp_size_t)(len / __gmpn_bases[base].chars_per_bit_exactly)) / GMP_NUMB_BITS + 2);
	if (sz > z->_mp_alloc)
		_mpz_realloc(z, sz);
	z->_mp_size = (int)mpn_set_str(z->_mp_d, (const unsigned char *)str, len, base);
}

bool Decimal::assign(LPCTSTR str)
{
	bool negative;
	if (*str == '-')
		negative = true, str++;
	else negative = false;
	//wcslen
	size_t len = _tcslen(str);
	size_t dot = 0, e_pos = 0, i;
	char *buf = (char *)malloc(len + 2), *p = buf;
	int base = 10;
	if (!buf)
		return false;
	e = 0;
	if (*str == '0') {
		if (str[1] == 'x' || str[1] == 'X')
			base = 16, str += 2;
		else if (str[1] == 'b' || str[1] == 'B')
			base = 2, str += 2;
	}
	for (i = 0; i < len; i++) {
		auto c = (unsigned char)str[i];
		if (__gmp_digit_value_tab[c] < base) {
			*p++ = __gmp_digit_value_tab[c];
			continue;
		}
		else if (c == '.' && dot == 0) {
			dot = i + 1;
			continue;
		}
		else if (base == 10 && (c == 'e' || c == 'E') && i) {
			e_pos = i++, e = 0;
			bool neg = false;
			if (str[i] == '-')
				neg = true, i++;
			auto p = str + i;
			for (auto c = *p; c >= '0' && c <= '9'; e = e * 10 + c - '0', c = *++p);
			if (neg)
				e = -e;
			if (dot)
				e -= e_pos - dot;
			if (!*p)
				break;
			e = 0;
		}
		free(buf);
		return false;
	}
	*p = 0;
	if (!e_pos && dot)
		e = dot - len;
	assign(buf, p - buf, base);
	if (negative)
		z->_mp_size = -z->_mp_size;
	free(buf);
	return true;
}

void Decimal::carry(bool ignore_integer, bool fix)
{
	if (!z->_mp_size) {
		e = 0;
		return;
	}
	if (z->_mp_size == 1) {
		auto &p = *z->_mp_d;
		if (!ignore_integer)
			while (!(p % 10))
				p /= 10, ++e;
		else if (e < 0)
			while (!(p % 10)) {
				p /= 10, ++e;
				if (!e)
					break;
			}
		if (!fix)
			return;
		auto c = (mpir_si)log10(double(p)) + 1 + sPrec;
		if (c > 0) {
			mp_limb_t t = 1;
			for (auto i = c; i--; t *= 10);
			p = (p + (t >> 1)) / t;
			if (!e)
				p *= t;
			else e += c;
		}
	}
	else if (fix || (z->_mp_d[0] & 1)) {
		auto sz = z->_mp_size < 0 ? -z->_mp_size : z->_mp_size;
		auto res_buf = (unsigned char *)malloc((size_t)(0.3010299956639811 * GMP_NUMB_BITS * sz + 3)), res_str = res_buf + 1;
		if (!res_buf)
			return;
		size_t str_size = mpn_get_str(res_str, 10, z->_mp_d, sz), raw_size = str_size, i;
		if (fix && (size_t)-sPrec < str_size) {
			e += str_size + sPrec;
			if (res_str[str_size = -sPrec] > 4) {
				res_buf[0] = 0;
				for (auto p = res_str + str_size - 1; ++(*p) == 10; *p-- = 0);
				if (res_buf[0])
					res_str--, str_size++;
			}
		}
		for (i = str_size - 1; i > 0 && !res_str[i]; i--);
		if (i = str_size - i - 1) {
			if ((e += i) > 0 && ignore_integer)
				i -= (size_t)e, e = 0;
			str_size -= i;
		}
		if (raw_size != str_size) {
			auto sz = (int)mpn_set_str(z->_mp_d, res_str, str_size, 10);
			z->_mp_size = z->_mp_size < 0 ? -sz : sz;
		}
		free(res_buf);
	}
}

bool Decimal::is_integer()
{
	if (e > 0) {
		mul_10exp(this, this, (mpir_ui)e), e = 0;
		return true;
	}
	else carry();
	return e == 0;
}

LPTSTR Decimal::to_string()
{
	static char ch[] = "0123456789";
	auto sz = z->_mp_size;
	bool zs = false, negative = false;
	if (sz < 0)
		sz = -sz, negative = true;
	auto res_buf = (unsigned char *)malloc((size_t)(0.3010299956639811 * GMP_NUMB_BITS * sz + 3));
	if (!res_buf)
		return nullptr;
	auto res_str = res_buf + 1;
	auto e = this->e;
	size_t str_size = mpn_get_str(res_str, 10, z->_mp_d, sz), i = 0;
	if (sOutputPrec > 0 && (size_t)sOutputPrec < str_size) {
		e += str_size - sOutputPrec;
		if (res_str[str_size = sOutputPrec] > 4) {
			res_buf[0] = 0;
			for (auto p = res_str + str_size - 1; ++(*p) == 10; *p-- = 0);
			if (res_buf[0])
				res_str--, str_size++;
		}
	}
	for (; --str_size > 1 && !res_str[str_size]; e++);
	str_size++;
	auto el = e == 0 ? 0 : (__int64)log10(double(e > 0 ? e : -e)) + 3;
	LPTSTR buf, p;

	if (e > 0) {
		p = buf = (LPTSTR)malloc(sizeof(TCHAR) * (size_t)(1 + str_size + (e < el ? e : el) + negative));
		if (negative)
			*p++ = '-';
		if (str_size + e > 20 && e > el) {
			*p++ = ch[res_str[i++]];
			*p++ = '.';
			zs = true;
		}
	}
	else if (e < 0) {
		el++;
		if ((size_t)-e < str_size) {
			p = buf = (LPTSTR)malloc(sizeof(TCHAR) * (2 + str_size + negative));
			if (negative)
				*p++ = '-';
			for (size_t n = str_size + e; i < n;)
				*p++ = ch[res_str[i++]];
			*p++ = '.';
		}
		else if (e < -18 && (size_t)(2ull - e) >(size_t)(str_size + el)) {
			p = buf = (LPTSTR)malloc(sizeof(TCHAR) * (size_t)(1 + el + str_size + negative));
			if (negative)
				*p++ = '-';
			*p++ = ch[res_str[i++]];
			*p++ = '.';
			zs = true;
		}
		else {
			p = buf = (LPTSTR)malloc(sizeof(TCHAR) * (3 - e + str_size + negative));
			if (negative)
				*p++ = '-';
			*p++ = '0';
			*p++ = '.';
			for (size_t n = -e - str_size; n > 0; n--)
				*p++ = '0';
		}
	}
	else {
		p = buf = (LPTSTR)malloc(sizeof(TCHAR) * (1 + str_size + negative));
		if (negative)
			*p++ = '-';
	}
	for (size_t n = str_size; i < n;)
		*p++ = ch[res_str[i++]];
	if (zs) {
		*p++ = 'e';
		if (e < 0)
			*p++ = '-', e = -e, el--;
		p += el - 2;
		*p = 0;
		while (e)
			*--p = (e % 10) + '0', e /= 10;
	}
	else {
		while (e > 0)
			*p++ = '0', e--;
		*p = 0;
	}
	free(res_buf);
	return buf;
}

void Decimal::mul_10exp(Decimal *v, Decimal *a, mpir_ui e)
{
	if (!a->z->_mp_size)
		v->z->_mp_size = 0;
	else if (e == 0)
		mpz_set(v->z, a->z);
	else if (a == v) {
		mpz_t t;
		mpz_init(t);
		mpz_ui_pow_ui(t, 10, e);
		mpz_mul(v->z, t, a->z);
		mpz_clear(t);
	}
	else {
		auto z = v->z;
		mpz_ui_pow_ui(z, 10, e);
		mpz_mul(z, z, a->z);
	}
}

void Decimal::add_or_sub(Decimal *v, Decimal *a, Decimal *b, int add)
{
	if (!b->z->_mp_size) {
		mpz_set(v->z, a->z);
		v->e = a->e;
		if (sPrec < 0)
			v->carry(true, true);
		return;
	}

	if (!a->z->_mp_size) {
		mpz_set(v->z, b->z);
		v->e = b->e;
		if (!add)
			v->z->_mp_size = -v->z->_mp_size;
		if (sPrec < 0)
			v->carry(true, true);
		return;
	}

	if (auto edif = a->e - b->e) {
		Decimal *h, *l, t;
		if (edif < 0)
			h = b, l = a, edif = -edif;
		else h = a, l = b;
		mul_10exp(&t, h, edif);
		if ((v->e = l->e) >= 0)
			mpz_swap(t.z, h->z), h->e = l->e;
		else {
			t.e = l->e;
			if (h == a)
				a = &t;
			else b = &t;
		}

		if (add)
			mpz_add(v->z, a->z, b->z);
		else mpz_sub(v->z, a->z, b->z);
	}
	else {
		v->e = a->e;
		if (add)
			mpz_add(v->z, a->z, b->z);
		else mpz_sub(v->z, a->z, b->z);
	}

	if (sPrec < 0)
		v->carry(true, true);
	else if (v->e < sCarryPrec || sPrec < 0)
		v->carry();
}

void Decimal::mul(Decimal *v, Decimal *a, Decimal *b)
{
	mpz_mul(v->z, a->z, b->z);
	if (sPrec < 0)
		v->carry(true, true);
	else if ((v->e = v->z->_mp_size ? a->e + b->e : 0) < sCarryPrec)
		v->carry();
}

int Decimal::div(Decimal *v, Decimal *a, Decimal *b, bool intdiv)
{
	auto dd = a->z, d = b->z;
	if (!b->z->_mp_size)
		return 0;
	b->carry(false);
	if (!dd->_mp_size) {
		v->z->_mp_d[0] = v->e = v->z->_mp_size = 0;
		return 1;
	}

	Decimal quot, rem;
	if (d->_mp_d[0] == 1 && (d->_mp_size == 1 || d->_mp_size == -1)) {
		if (intdiv && a->e < 0) {
			mpz_ui_pow_ui(quot.z, 10, -a->e);
			mpz_tdiv_qr(v->z, rem.z, a->z, quot.z);
			v->e = 0;
		}
		else a->copy_to(v);
		if (d->_mp_size < 0)
			v->z->_mp_size = -v->z->_mp_size;
		if (sPrec < 0)
			v->carry(true, true);
		return 1;
	}

	auto aw = mpz_sizeinbase(dd, 10), bw = mpz_sizeinbase(d, 10);
	mpir_si exp = 0;
	int eq;
	v->e = a->e - b->e;
	if (exp = bw - aw) {
		Decimal *aa, *bb;
		if (exp > 0)
			mul_10exp(aa = &quot, a, exp), bb = b;
		else
			mul_10exp(bb = &quot, b, -exp), aa = a;
		eq = mpz_cmp(aa->z, bb->z);
	}
	else eq = mpz_cmp(a->z, b->z);
	if (eq == 0) {
		if (exp >= 0) {
			v->z->_mp_d[0] = v->z->_mp_size = 1;
			v->e = -exp;
			return 1;
		}
		else {
			mpz_ui_pow_ui(v->z, 10, -exp);
			v->e = 0;
			return 1;
		}
	}

	exp += (sPrec < 0 ? -sPrec : sPrec) - (eq > 0);
	if (eq > 0 || aw > bw) {
		mpz_tdiv_qr(quot.z, rem.z, dd, d);
		if (!rem.z->_mp_size || exp <= 0)
			goto divend;
	}

	if (exp > 0) {
		mul_10exp(&quot, a, exp);
		mpz_tdiv_qr(quot.z, rem.z, quot.z, d), v->e -= exp;
	}
	else
		mpz_tdiv_qr(quot.z, rem.z, dd, d);

divend:
	mpz_swap(quot.z, v->z);

	if (exp < 0) {
		mpz_ui_pow_ui(d = quot.z, 10, -exp);
		mpz_tdiv_qr(v->z, rem.z, v->z, d);
		v->e -= exp;
	}
	if (intdiv) {
		if (v->e < 0) {
			mpz_ui_pow_ui(quot.z, 10, -v->e);
			mpz_tdiv_qr(v->z, rem.z, v->z, quot.z);
			v->e = 0;
		}
	}
	else if (rem.z->_mp_size) {
		mpz_mul_2exp(rem.z, rem.z, 1);
		if (mpz_cmp(rem.z, d) >= 0)
			if (v->z->_mp_size > 0)
				mpz_add_ui(v->z, v->z, 1);
			else
				mpz_sub_ui(v->z, v->z, 1);
	}
	if (v->e < sCarryPrec)
		v->carry();
	return 1;
}

void Decimal::SetPrecision(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	if (!TokenIsNumeric(*aParam[1]))
		_f_throw_param(1, _T("Integer"));
	aResultToken.SetValue(sPrec);
	auto v = (mpir_si)TokenToInt64(*aParam[1]);
	sPrec = v ? v : 20;
	if (aParamCount > 2 && aParam[2]->symbol != SYM_MISSING)
		if ((sOutputPrec = (mpir_si)TokenToInt64(*aParam[2])) < 0)
			sOutputPrec = -sOutputPrec;
}

Decimal *Decimal::Create(ExprTokenType *aToken)
{
	auto obj = new Decimal;
	if (aToken) {
		switch (aToken->symbol)
		{
		case SYM_INTEGER:obj->assign(aToken->value_int64); break;
		case SYM_FLOAT:obj->assign(aToken->value_double); break;
		case SYM_STRING:
			if (obj->assign(aToken->marker))
				break;

		default:
			if (auto t = ToDecimal(*aToken)) {
				obj->e = t->e;
				mpz_set(obj->z, t->z);
			}
			else {
				delete obj;
				obj = nullptr;
			}
			break;
		}
	}
	return obj;
}

void Decimal::Create(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	ExprTokenType val, *aToken = aParam[1];
	if (aToken->symbol == SYM_VAR)
		aToken->var->ToTokenSkipAddRef(val), aToken = &val;
	if (auto obj = Decimal::Create(aParam[1]))
		aResultToken.SetValue(obj);
	else
		_f_throw_param(0, _T("Number"));
}

void Decimal::Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	switch (aID)
	{
	case M_ToString: {
		aResultToken.symbol = SYM_STRING;
		aResultToken.mem_to_free = aResultToken.marker = to_string();
		break;
	}
	default:
		break;
	}
}

int Decimal::Eval(ExprTokenType &this_token, ExprTokenType *right_token)
{
	Decimal tmp, *right = &tmp, *left = this, *ret = nullptr;
	SymbolType ret_symbol = SYM_OBJECT;
	if (right_token) {
		ExprTokenType tk;
		if (right_token->symbol == SYM_VAR)
			right_token->var->ToTokenSkipAddRef(tk), right_token = &tk;
		if ((this_token.symbol >= SYM_BITSHIFTLEFT && this_token.symbol <= SYM_BITSHIFTRIGHT_LOGICAL) || this_token.symbol == SYM_POWER) {
			bool neg = false;
			mpir_ui n;
			if (right_token->symbol == SYM_STRING) {
				auto t = IsNumeric(right_token->marker, true, false, true);
				if (t == PURE_INTEGER)
					tk.SetValue(TokenToInt64(*right_token)), right_token = &tk;
				else return t ? -2 : -1;
			}
			if (right_token->symbol == SYM_INTEGER) {
				if (right_token->value_int64 < 0)
					if (this_token.symbol == SYM_POWER)
						neg = true, n = (mpir_ui) - right_token->value_int64;
					else return -2;
				else n = (mpir_ui)right_token->value_int64;
			}
			else if (right_token->symbol == SYM_FLOAT)
				return -2;
			else if (auto dec = ToDecimal(*right_token)) {
				if (!dec->is_integer() || dec->z->_mp_size != 1)
					return -2;
				n = dec->z->_mp_d[0];
			}
			else return -1;

			if (this_token.symbol == SYM_POWER) {
				ret = new Decimal;
				mpz_pow_ui(ret->z, z, n), ret->e = e * n;
				if (neg) {
					tmp.assign(1LL);
					div(ret, &tmp, ret);
				}
			}
			else {
				if (!is_integer())
					return -2;
				ret = new Decimal;
				if (this_token.symbol == SYM_BITSHIFTLEFT)
					mpz_mul_2exp(ret->z, z, n);
				else
					mpz_div_2exp(ret->z, z, n);
			}
		}
		else {
			mpir_si diff;
			if (right_token->symbol == SYM_STRING) {
				if (!tmp.assign(right_token->marker))
					return -1;
			}
			else if (right_token->symbol == SYM_INTEGER)
				tmp.assign(right_token->value_int64);
			else if (right_token->symbol == SYM_FLOAT)
				tmp.assign(right_token->value_double);
			else if (auto t = ToDecimal(right_token->object))
				right = t;
			else return -1;

			ret_symbol = SYM_OBJECT;
			ret = new Decimal;
			switch (this_token.symbol)
			{
			case SYM_ADD: add_or_sub(ret, this, right); break;
			case SYM_SUBTRACT: add_or_sub(ret, this, right, false); break;
			case SYM_MULTIPLY: mul(ret, this, right); break;
			case SYM_DIVIDE:
			case SYM_INTEGERDIVIDE:
				if (div(ret, this, right, this_token.symbol == SYM_INTEGERDIVIDE) == 0) {
					delete ret;
					return 0;
				}
				break;

			case SYM_BITAND:
			case SYM_BITOR:
			case SYM_BITXOR:
				if (!left->is_integer() || !right->is_integer()) {
					delete ret;
					return -2;
				}
				if (this_token.symbol == SYM_BITAND)
					mpz_and(ret->z, left->z, right->z);
				else if (this_token.symbol == SYM_BITOR)
					mpz_ior(ret->z, left->z, right->z);
				else
					mpz_xor(ret->z, left->z, right->z);
				break;

			default:
				delete ret;
				ret_symbol = SYM_INTEGER;
				diff = e - right->e;
				if (diff < 0)
					mul_10exp(&tmp, right, -diff), right = &tmp;
				else if (diff > 0)
					mul_10exp(&tmp, left, diff), left = &tmp;
				diff = mpz_cmp(left->z, right->z);
				switch (this_token.symbol)
				{
				case SYM_EQUAL:
				case SYM_EQUALCASE: this_token.value_int64 = diff == 0; break;
				case SYM_NOTEQUAL:
				case SYM_NOTEQUALCASE: this_token.value_int64 = diff != 0; break;
				case SYM_GT: this_token.value_int64 = diff > 0; break;
				case SYM_LT: this_token.value_int64 = diff < 0; break;
				case SYM_GTOE: this_token.value_int64 = diff >= 0; break;
				case SYM_LTOE: this_token.value_int64 = diff <= 0; break;
				default:
					return -1;
				}
			}
		}
		if ((this_token.symbol = ret_symbol) == SYM_OBJECT)
			this_token.SetValue(ret);
		return 1;
	}

	if (this_token.symbol == SYM_POSITIVE)
		AddRef(), ret = this;
	else {
		ret = new Decimal;
		tmp.assign(1LL);
		switch (this_token.symbol)
		{
		case SYM_NEGATIVE:       copy_to(ret), ret->z->_mp_size = -z->_mp_size; break;
		case SYM_POST_INCREMENT: copy_to(ret), add_or_sub(this, this, &tmp); break;
		case SYM_POST_DECREMENT: copy_to(ret), add_or_sub(this, this, &tmp, 0); break;
		case SYM_PRE_INCREMENT:  add_or_sub(this, this, &tmp), copy_to(ret); break;
		case SYM_PRE_DECREMENT:  add_or_sub(this, this, &tmp, 0), copy_to(ret); break;
		//case SYM_BITNOT:
		default:
			delete ret;
			return -1;
		}
	}
	this_token.SetValue(ret);
	return 1;
}

Decimal *Decimal::ToDecimal(ExprTokenType &aToken)
{
	if (aToken.symbol == SYM_OBJECT)
		return ToDecimal(aToken.object);
	if (aToken.symbol == SYM_VAR)
		if (auto obj = aToken.var->ToObject())
			return ToDecimal(obj);
	return nullptr;
}

ResultType Decimal::ToToken(ExprTokenType &aToken)
{
	if (is_integer())
		aToken.SetValue(mpz_get_sx(z));
	else {
		auto t = sOutputPrec;
		sOutputPrec = 17;
		auto str = to_string();
		sOutputPrec = t;
		aToken.SetValue(ATOF(str));
		free(str);
	}
	return OK;
}

void *Decimal::sVTable = getVTable();
thread_local Object *Decimal::sPrototype;
thread_local mpir_si Decimal::sPrec = 20;
thread_local mpir_si Decimal::sOutputPrec = 0;

#endif // ENABLE_DECIMAL
