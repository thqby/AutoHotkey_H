#include "stdafx.h" // pre-compiled headers
#include "globaldata.h" // for access to many global vars
#include "script_func_impl.h"

BIF_DECL(BIF_GetVar)
{
	Var *var = nullptr;
	if (aParam[0]->symbol == SYM_VAR)
		var = aParam[0]->var;
	else var = g_script->FindVar(TokenToString(*aParam[0]), 0, FINDVAR_DEFAULT);
	if (!var)
		_o_throw(ERR_DYNAMIC_NOT_FOUND);
	if (var->IsAlias() && ParamIndexToOptionalBOOL(1, TRUE))
		var = var->ResolveAlias();
	_o_return((__int64)var);
}


BIF_DECL(BIF_Alias)
{
	ExprTokenType aParam0 = *aParam[0];
	if (aParam0.symbol == SYM_STRING)
		if (auto v = g_script->FindVar(aParam0.marker, aParam0.marker_length, FINDVAR_DEFAULT))
			aParam0.SetVar(v);
		else _o_throw(ERR_DYNAMIC_NOT_FOUND);
	if (aParam0.symbol != SYM_VAR)
		_o_throw_param(0);
	Var &var = *aParam0.var, *target = nullptr;
	if (aParamCount < 2 || aParam[1]->symbol == SYM_MISSING)
	{
		var.Free(VAR_ALWAYS_FREE | VAR_CLEAR_ALIASES | VAR_REQUIRE_INIT);
		return;
	}
	if (aParam[1]->symbol == SYM_VAR)
		target = aParam[1]->var;
	else
		target = (Var *)TokenToInt64(*aParam[1]);
	if (target < (void *)65536 || target->ResolveAlias() == &var)
		_o_throw_param(1);
	var.Free(VAR_ALWAYS_FREE | VAR_CLEAR_ALIASES | VAR_REQUIRE_INIT);
	var.SetAliasDirect(target);
}