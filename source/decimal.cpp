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
	char *buf = (char *)_malloca(len + 2), *p = buf;
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
		_freea(buf);
		return false;
	}
	*p = 0;
	if (!e_pos && dot)
		e = dot - len;
	for (; --p > buf && !*p; e++);
	p++;
	assign(buf, p - buf, base);
	if (!z->_mp_size)
		e = 0;
	else if (negative)
		z->_mp_size = -z->_mp_size;
	_freea(buf);
	return true;
}

void Decimal::carry(bool ignore_integer, bool fix)
{
	if (!z->_mp_size) {
		e = 0;
		return;
	}
	auto prec = sPrec < 0 ? -sPrec : sPrec;
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
		auto c = (mpir_si)log10(double(p)) + 1 - prec;
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
		auto res_buf = (unsigned char *)_malloca(size_t(3 + sz * (__gmpn_bases[10].chars_per_bit_exactly * GMP_NUMB_BITS + SIZEOF_MP_LIMB_T)));
		if (!res_buf)
			return;
		auto res_str = res_buf + 1 + sz * SIZEOF_MP_LIMB_T, xp = res_buf;
		memcpy(xp, z->_mp_d, sz * SIZEOF_MP_LIMB_T);
		size_t str_size = mpn_get_str(res_str, 10, (mp_limb_t*)xp, sz), raw_size = str_size, i;
		if (fix && (size_t)prec < str_size) {
			e += str_size - prec;
			if (res_str[str_size = prec] > 4) {
				res_str[-1] = 0;
				for (auto p = res_str + str_size - 1; ++(*p) == 10; *p-- = 0);
				if (res_str[-1])
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
		_freea(res_buf);
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

LPTSTR Decimal::to_string(size_t *aLength, mpir_si *prec)
{
	static char ch[] = "0123456789";
	auto sz = z->_mp_size;
	bool zs = false, negative = false;
	if (sz < 0)
		sz = -sz, negative = true;
	else if (sz == 0)
		e = 0;
	auto res_buf = (unsigned char *)_malloca(size_t(2 + sz * (__gmpn_bases[10].chars_per_bit_exactly * GMP_NUMB_BITS + SIZEOF_MP_LIMB_T)));
	if (!res_buf)
		return nullptr;
	auto e = this->e;
	auto res_str = res_buf + sz * SIZEOF_MP_LIMB_T, xp = res_buf;
	memcpy(xp, z->_mp_d, sz * SIZEOF_MP_LIMB_T);
	size_t str_size = mpn_get_str(res_str, 10, (mp_limb_t*)xp, sz), i = 0;
	mpir_si outputprec = prec ? *prec : sOutputPrec, ws = 0;
	LPTSTR buf, p;
	if (outputprec > 0) {
		ws = outputprec;
		mpir_si d = (mpir_si)str_size + e, l;
		if (d > 0)
			l = outputprec + d + 3 + negative;
		else
			l = outputprec + 4 + negative;
		outputprec = -outputprec - d;
		p = buf = (LPTSTR)malloc(sizeof(TCHAR) * l);
		if (outputprec >= 0) {
			if (negative && !outputprec && res_str[0] > 4)
				*p++ = '-';
			*p++ = '0', *p++ = '.';
			for (mpir_si i = 0; i < ws; i++)
				*p++ = '0';
			if (!outputprec && res_str[0] > 4)
				p[-1]++;
			*p = 0, ws = -1;
		}
	}
	if (outputprec < 0) {
		outputprec = -outputprec;
		if ((size_t)outputprec < str_size) {
			e += str_size - outputprec;
			if (res_str[str_size = (size_t)outputprec] > 4) {
				for (auto p = res_str + str_size - 1; ++(*p) == 10 && p > res_str; *p-- = 0);
				if (res_str[0] == 10)
					e += str_size, str_size = res_str[0] = 1;
			}
		}
	}
	if (ws) {
		if (ws > 0) {
			__int64 d = (__int64)str_size + e, i = 0;
			if (negative)
				*p++ = '-';
			if (d > 0) {
				auto dl = min(d, (__int64)str_size);
				for (; i < dl; i++)
					*p++ = ch[res_str[i]];
				for (; i < d; i++)
					*p++ = '0';
				*p++ = '.';
				i = 0;
				auto src = res_str + dl;
				dl = min(ws, __int64(str_size - dl));
				for (; i < dl; i++)
					*p++ = ch[*src++];
			}
			else {
				*p++ = '0', *p++ = '.';
				for (d = -d; i < d; i++)
					*p++ = '0';
				auto dl = min(ws - d, (__int64)str_size) + i;
				for (auto src = res_str; i < dl; i++)
					*p++ = ch[*src++];
			}
			for (; i < ws; i++)
				*p++ = '0';
			*p = 0;
		}
	}
	else {
		for (; --str_size > 0 && !res_str[str_size]; e++);
		str_size++;

		auto el = e == 0 ? 0 : (__int64)log10(double(e > 0 ? e : -e)) + 3;
		size_t yxsize = outputprec ? outputprec + 1 : str_size;
		bool has_dot = false;
		if (e > 0) {
			if (outputprec && str_size + e > size_t(outputprec) || str_size + e > 20 && e > el) {
				p = buf = (LPTSTR)malloc(sizeof(TCHAR) * size_t(1 + yxsize + el + negative));
				if (negative)
					*p++ = '-';
				*p++ = ch[res_str[i++]];
				*p++ = '.', has_dot = true;
				zs = true;
			}
			else {
				p = buf = (LPTSTR)malloc(sizeof(TCHAR) * size_t(1 + yxsize + e + negative));
				if (negative)
					*p++ = '-';
			}
		}
		else if (e < 0) {
			el++;
			if ((size_t)-e < str_size) {
				p = buf = (LPTSTR)malloc(sizeof(TCHAR) * (2 + yxsize + negative));
				if (negative)
					*p++ = '-';
				for (size_t n = str_size + e; i < n;)
					*p++ = ch[res_str[i++]];
				*p++ = '.', has_dot = true;
			}
			else if (e < -18 && size_t(2ull - e) > size_t(str_size + el)) {
				p = buf = (LPTSTR)malloc(sizeof(TCHAR) * size_t(1 + el + yxsize + negative));
				if (negative)
					*p++ = '-';
				*p++ = ch[res_str[i++]];
				*p++ = '.', has_dot = true;
				zs = true;
			}
			else {
				p = buf = (LPTSTR)malloc(sizeof(TCHAR) * (3 - e + yxsize + negative));
				if (negative)
					*p++ = '-';
				*p++ = '0';
				*p++ = '.', has_dot = true;
				for (size_t n = -e - str_size; n > 0; n--)
					*p++ = '0';
			}
		}
		else {
			p = buf = (LPTSTR)malloc(sizeof(TCHAR) * (1 + yxsize + negative));
			if (negative)
				*p++ = '-';
		}
		for (size_t n = str_size; i < n;)
			*p++ = ch[res_str[i++]];
		if (i < size_t(outputprec)) {
			auto end = size_t(outputprec);
			if (str_size + e == end)
				while (e > 0)
					*p++ = '0', e--, i++;
			if (i < end) {
				if (!has_dot)
					*p++ = '.';
				for (; i < end; i++)
					*p++ = '0';
			}
		}
		if (zs) {
			if (p[-1] == '.')
				p--, has_dot = false;
			*p++ = 'e';
			if (e < 0)
				*p++ = '-', e = -e, el--;
			p += el - 2;
			*p = 0;
			for (auto pp = p; e;)
				*--pp = (e % 10) + '0', e /= 10;
		}
		else {
			while (e > 0)
				*p++ = '0', e--;
			*p = 0;
		}
	}
	_freea(res_buf);
	if (aLength)
		*aLength = p - buf;
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
	if (intdiv) {
		auto ediff = a->e - b->e;
		if (ediff < 0) {
			mpz_ui_pow_ui(quot.z, 10, -ediff);
			mpz_tdiv_qr(dd = v->z, rem.z, a->z, quot.z);
		}
		else if (ediff > 0) {
			mpz_ui_pow_ui(quot.z, 10, ediff);
			mpz_mul(dd = v->z, a->z, quot.z);
		}
		mpz_tdiv_qr(v->z, rem.z, dd, d);
		v->e = 0;
		if (sPrec < 0)
			v->carry(true, true);
		return 1;
	}
	if (d->_mp_d[0] == 1 && (d->_mp_size == 1 || d->_mp_size == -1)) {
		v->e = a->e - b->e;
		mpz_set(v->z, a->z);
		v->carry(true, true);
		if (d->_mp_size < 0)
			v->z->_mp_size = -v->z->_mp_size;
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
		eq = mpz_cmpabs(aa->z, bb->z);
	}
	else eq = mpz_cmpabs(a->z, b->z);
	if (eq == 0) {
		v->z->_mp_d[0] = 1;
		v->z->_mp_size = (dd->_mp_size ^ d->_mp_size) < 0 ? -1 : 1;
		v->e -= exp;
		return 1;
	}

	exp += (sPrec < 0 ? -sPrec : sPrec) - (eq > 0);
	if (eq > 0 && exp <= 0 || aw > bw) {
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
	if (rem.z->_mp_size) {
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
	aResultToken.SetValue(sPrec);
	if (aParamCount < 2)
		return;
	if (aParam[1]->symbol != SYM_MISSING) {
		if (!TokenIsNumeric(*aParam[1]))
			_f_throw_param(1, _T("Integer"));
		auto v = (mpir_si)TokenToInt64(*aParam[1]);
		sPrec = v ? v : 20;
		sCarryPrec = -max((v < 0 ? -v : v) >> 1, 5);
	}
	if (aParamCount > 2 && aParam[2]->symbol != SYM_MISSING)
		sOutputPrec = (mpir_si)TokenToInt64(*aParam[2]);
}

bool Decimal::Assign(ExprTokenType *aToken)
{
	ExprTokenType tmp;
	if (!aToken)
		aToken = &tmp, tmp.SetValue(0);
	else if (aToken->symbol == SYM_VAR)
		aToken->var->ToTokenSkipAddRef(tmp), aToken = &tmp;
	switch (aToken->symbol)
	{
	case SYM_STRING:
		return assign(aToken->marker);
	case SYM_INTEGER:
		assign(aToken->value_int64);
		return true;
	case SYM_FLOAT:
		assign(aToken->value_double);
		return true;
	case SYM_OBJECT:
		if (auto obj = ToDecimal(aToken->object)) {
			obj->copy_to(this);
			return true;
		}
	default:
		return false;
	}
}

Decimal *Decimal::Create(ExprTokenType *aToken)
{
	auto obj = new Decimal;
	if (obj && !obj->Assign(aToken)) {
		delete obj;
		return nullptr;
	}
	return obj;
}

void Decimal::Create(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	++aParam;
	if (auto obj = Decimal::Create(aParam[0]))
		_o_return(obj);
	_o_throw_param(0, _T("Number"));
}

void Decimal::Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	switch (aID)
	{
	case M_ToString: {
		mpir_si outputprec, *p = nullptr;
		if (aParamCount)
			outputprec = (mpir_si)TokenToInt64(*aParam[0]), p = &outputprec;
		if (auto str = to_string(&aResultToken.marker_length, p)) {
			aResultToken.symbol = SYM_STRING;
			aResultToken.mem_to_free = aResultToken.marker = str;
			break;
		}
		_o_throw_oom;
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
						neg = true, n = (mpir_ui)-right_token->value_int64;
					else return -2;
				else n = (mpir_ui)right_token->value_int64;
			}
			else if (right_token->symbol == SYM_FLOAT)
				return -2;
			else if (auto dec = ToDecimal(*right_token)) {
				if (!dec->is_integer())
					return -2;
				else if (dec->z->_mp_size == 0)
					dec->z->_mp_d[0] = 0;
				else if (dec->z->_mp_size != 1)
					if (dec->z->_mp_size == -1 && this_token.symbol == SYM_POWER)
						neg = true;
					else return -2;
				n = dec->z->_mp_d[0];
			}
			else return -1;

			if (this_token.symbol == SYM_POWER) {
				if (!n && !z->_mp_size)
					return -2;
				ret = new Decimal;
				if (n) {
					mpz_pow_ui(ret->z, z, n), ret->e = e * n;
					if (neg) {
						tmp.assign(1LL);
						div(ret, &tmp, ret);
					}
				}
				else ret->assign(1LL);
			}
			else {
				if (!is_integer())
					return -2;
				ret = new Decimal;
				if (!n)
					copy_to(ret);
				else if (this_token.symbol == SYM_BITSHIFTLEFT)
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
			case SYM_SUBTRACT: add_or_sub(ret, this, right, 0); break;
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

	tmp.assign(1LL);
	switch (this_token.symbol)
	{
	case SYM_POSITIVE:       ret = Clone(); if (sPrec < 0) ret->carry(); break;
	case SYM_NEGATIVE:       ret = Clone(), ret->z->_mp_size = -z->_mp_size; if (sPrec < 0) ret->carry(); break;
	case SYM_POST_INCREMENT: ret = Clone(), add_or_sub(this, this, &tmp); break;
	case SYM_POST_DECREMENT: ret = Clone(), add_or_sub(this, this, &tmp, 0); break;
	case SYM_PRE_INCREMENT:  AddRef(), add_or_sub(ret = this, this, &tmp); break;
	case SYM_PRE_DECREMENT:  AddRef(), add_or_sub(ret = this, this, &tmp, 0); break;
	//case SYM_BITNOT:
	default:
		return -1;
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
	if (!z->_mp_size) {
		aToken.SetValue(0);
		return OK;
	}

	if (is_integer() && size_t(z->_mp_size < 0 ? -z->_mp_size : z->_mp_size) * SIZEOF_MP_LIMB_T <= sizeof(__int64)) {
		auto sz = z->_mp_size < 0 ? -z->_mp_size : z->_mp_size;
#if (SIZEOF_MP_LIMB_T == 4)
		if (sz == 1) {
			aToken.SetValue(sz < 0 ? -__int64(z->_mp_d[0]) : __int64(z->_mp_d[0]));
			return OK;
		}
#endif // (SIZEOF_MP_LIMB_T == 4)
		if (sz * SIZEOF_MP_LIMB_T == 8) {
			auto v = *(unsigned __int64 *)z->_mp_d;
			if (v <= 9223372036854775807ui64 || (sz < 0 && v == 9223372036854775808ui64)) {
				aToken.SetValue(sz < 0 ? -__int64(v - 1) - 1 : __int64(v));
				return OK;
			}
		}
	}

	mpir_si prec = 17;
	auto str = to_string(nullptr, &prec);
	aToken.SetValue(ATOF(str));
	free(str);
	return OK;
}


void *Decimal::sVTable = getVTable();
thread_local Object *Decimal::sPrototype;
thread_local mpir_si Decimal::sCarryPrec = -10;
thread_local mpir_si Decimal::sPrec = 20;
thread_local mpir_si Decimal::sOutputPrec = 0;

#endif // ENABLE_DECIMAL
