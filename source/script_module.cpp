#include "stdafx.h"
#include "script.h"
#include "globaldata.h"


Object *ScriptModule::sPrototype;


ResultType ScriptModule::Invoke(IObject_Invoke_PARAMS_DECL)
{
	auto var = aName ? mVars.Find(aName) : nullptr;
	if (!var && mIsBuiltinModule)
		// This is a slight hack to support built-in vars which haven't been referenced directly.
		var = g_script.FindOrAddBuiltInVar(aName, true, nullptr);
	if (!var || !var->IsExported())
		return ObjectBase::Invoke(IObject_Invoke_PARAMS);

	if (IS_INVOKE_SET && aParamCount == 1)
		return var->Assign(*aParam[0]);

	var->Get(aResultToken);
	if (aResultToken.Exited() || aParamCount == 0 && IS_INVOKE_GET)
		return aResultToken.Result();

	return Object::ApplyParams(aResultToken, aFlags, aParam, aParamCount);
}


ResultType Script::ParseModuleDirective(LPCTSTR aName)
{
	int at;
	if (mModules.Find(aName, &at))
		return ScriptError(ERR_DUPLICATE_DECLARATION, aName);
	// TODO: Validate module names.
	aName = SimpleHeap::Alloc(aName);
	auto mod = new ScriptModule(aName);
	// Let any previous #Warn settings carry over from the previous module, by default.
	mod->Warn_LocalSameAsGlobal = mCurrentModule->Warn_LocalSameAsGlobal;
	mod->Warn_Unreachable = mCurrentModule->Warn_Unreachable;
	mod->Warn_VarUnset = mCurrentModule->Warn_VarUnset;
	if (!mModules.Insert(mod, at))
		return MemoryError();
	CloseCurrentModule();
	mCurrentModule = mod;
	return CONDITION_TRUE;
}


ResultType Script::ParseImportStatement(LPTSTR aBuf)
{
	aBuf = omit_leading_whitespace(aBuf);
	auto imp = new ScriptImport();
	imp->names = SimpleHeap::Alloc(aBuf);
	imp->mod = nullptr;
	imp->line_number = mCombinedLineNumber;
	imp->file_index = mCurrFileIndex;
	imp->next = mCurrentModule->mImports;
	mCurrentModule->mImports = imp;
	return OK;
}


Var *Script::FindImportedVar(LPCTSTR aVarName)
{
	for (auto imp = CurrentModule()->mImports; imp; imp = imp->next)
	{
		if (*imp->names == '*')
		{
			auto var = imp->mod->mVars.Find(aVarName);
			if (var && var->IsExported())
				return var;
		}
	}
	return nullptr;
}


// Add a new Var to the current module.
// Raises an error if a conflicting declaration exists.
// May use an existing Var if not previously marked as declared, such as if created by Export.
// Caller provides persistent memory for aVarName.
Var *Script::AddNewImportVar(LPTSTR aVarName)
{
	int at;
	auto var = mCurrentModule->mVars.Find(aVarName, &at);
	if (var)
	{
		// Only declared, assigned or exported variables should exist at this point.
		if (var->IsDeclared() || var->IsAssignedSomewhere())
		{
			ConflictingDeclarationError(_T("Import"), var);
			return nullptr;
		}
		var->Scope() |= VAR_DECLARED;
	}
	else
		var = new Var(aVarName, VAR_DECLARE_GLOBAL);
	if (!mCurrentModule->mVars.Insert(var, at))
	{
		delete var;
		MemoryError();
		return nullptr;
	}
	return var;
}


ResultType Script::ResolveImports()
{
	for (mCurrentModule = mLastModule; mCurrentModule; mCurrentModule = mCurrentModule->mPrev)
	{
		for (auto imp = mCurrentModule->mImports; imp; imp = imp->next)
		{
			if (!imp->mod && !ResolveImports(*imp))
				return FAIL;
		}
	}
	return OK;
}


ResultType Script::ResolveImports(ScriptImport &imp)
{
	// Set early for code size, in case of error.
	mCurrLine = nullptr;
	mCombinedLineNumber = imp.line_number;
	mCurrFileIndex = imp.file_index;

	LPTSTR cp = imp.names, mod_name, var_name = nullptr;
	if (*cp == '{' || *cp == '*')
	{
		if (*cp == '{')
			cp = _tcschr(cp, '}'); // Should always be found due to GetLineContExpr().
		cp = omit_leading_whitespace(cp + 1);
		if (_tcsnicmp(cp, _T("From"), 4) || !IS_SPACE_OR_TAB(cp[4]))
			return ScriptError(_T("Invalid import"), imp.names);
		cp = omit_leading_whitespace(cp + 5);
		mod_name = cp;
	}
	else
	{
		var_name = mod_name = cp;
		while (*cp && !IS_SPACE_OR_TAB(*cp)) ++cp;
		if (*cp)
		{
			*cp = '\0';
			cp = omit_leading_whitespace(cp + 1);
			if (_tcsnicmp(cp, _T("as"), 2) || !IS_SPACE_OR_TAB(cp[2]))
				return ScriptError(_T("Invalid import"), imp.names);
			var_name = omit_leading_whitespace(cp + 3);
		}
	}

	int at;
	if (  !(imp.mod = mModules.Find(mod_name, &at))  )
	{
		if (auto path = FindLibraryFile(mod_name, _tcslen(mod_name), true))
		{
			auto cur_mod = mCurrentModule;
			auto last_mod = mLastModule;
			mLastModule = nullptr; // Start a new chain.
			imp.mod = mCurrentModule = new ScriptModule(mod_name);
			if (!mModules.Insert(imp.mod, at))
				return MemoryError();
			if (!LoadIncludedFile(path, false, false))
				return FAIL;
			if (!CloseCurrentModule())
				return FAIL;
			if (!ResolveImports()) // Resolve imports in all modules that were just included.
				return FAIL;
			imp.mod->mPrev = last_mod; // Join to previous chain.
			mCurrentModule = cur_mod;
		}
		else
			return ScriptError(_T("Module not found"), mod_name);
	}

	if (var_name)
	{
		// Do not reuse mSelf or a previous Var created by an import even if mod_name == var_name,
		// since the exported status of the Var (VAR_EXPORTED) shouldn't propagate between modules.
		auto var = AddNewImportVar(var_name);
		if (!var)
			return FAIL;
		if (imp.mod->mSelf) // Default export.
		{
			// For code size, aliasing is used even for constants.  For non-dynamic references to
			// constants, PreparseVarRefs() eliminates both the alias and the var reference itself.
			var->UpdateAlias(imp.mod->mSelf);
		}
		else
		{
			var->Assign(imp.mod);
			var->MakeReadOnly();
		}
	}

	if (*imp.names == '{')
	{
		for (cp = imp.names + 1; *(cp = omit_leading_whitespace(cp)) != '}'; ++cp)
		{
			auto c = *(cp = find_identifier_end(var_name = mod_name = cp));
			*cp = '\0';
			while (IS_SPACE_OR_TAB(c)) c = *++cp; // Find next non-whitespace.
			auto exported = imp.mod->mVars.Find(mod_name);
			if (!exported)
				return ScriptError(_T("No such export"), mod_name);
			if (!_tcsnicmp(cp, _T("as"), 2) && IS_SPACE_OR_TAB(cp[2]))
			{
				var_name = omit_leading_whitespace(cp + 3);
				cp = find_identifier_end(var_name);
				c = *cp;
				*cp = '\0';
				while (IS_SPACE_OR_TAB(c)) c = *++cp; // Find next non-whitespace.
			}
			if (!(c == ',' || c == '}'))
			{
				*cp = c;
				return ScriptError(_T("Invalid import"), cp);
			}
			auto imported = AddNewImportVar(var_name);
			if (!imported)
				return FAIL;
			imported->UpdateAlias(exported); // See the other UpdateAlias call above for comments.
			if (c == '}')
				break;
		}
	}

	return OK;
}


ResultType Script::CloseCurrentModule()
{
	// Terminate each module with RETURN so that all labels have a target and
	// all control flow statements that need it have a non-null mRelatedLine.
	if (!AddLine(ACT_EXIT))
		return FAIL;

	mCurrentModule->mPrev = mLastModule;
	mLastModule = mCurrentModule;

	ASSERT(!mCurrentModule->mFirstLine || mCurrentModule->mFirstLine == mFirstLine);
	mCurrentModule->mFirstLine = mFirstLine;
	mFirstLine = nullptr; // Start a new linked list of lines.
	mLastLine = nullptr;
	mLastLabel = nullptr; // Start a new linked list of labels.
	return OK;
}


bool ScriptModule::HasFileIndex(FileIndexType aFile)
{
	for (int i = 0; i < mFilesCount; ++i)
		if (mFiles[i] == aFile)
			return true;
	return false;
}


ResultType ScriptModule::AddFileIndex(FileIndexType aFile)
{
	if (mFilesCount == mFilesCountMax)
	{
		auto new_size = mFilesCount ? mFilesCount * 2 : 8;
		auto new_files = (FileIndexType*)realloc(mFiles, new_size * sizeof(FileIndexType));
		if (!new_files)
			return MemoryError();
		mFiles = new_files;
		mFilesCountMax = new_size;
	}
	mFiles[mFilesCount++] = aFile;
	return OK;
}


IObject *ScriptModule::FindGlobalObject(LPCTSTR aName)
{
	if (auto var = mVars.Find(aName))
	{
		ASSERT(!var->IsVirtual());
		return var->ToObject();
	}
	return nullptr;
}
