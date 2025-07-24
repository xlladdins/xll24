// register.h - Excel function and macro registration.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include "args.h"

namespace xll {

	// Define hidden name procedure pointing to array of arguments
	inline int RegProcedure(const Args& args)
	{
		XLOPER12 res, multi = args.asMulti();
		int ret = Excel12(xlcDefineName, &res, 5,
			(XLOPER12*)&args.procedure,
			&multi, // refers to
			&Nil,   // macro type
			&Nil,   // shortcut
			&True); // hidden

		auto rrr = Excel(xlcDefineName, args.procedure, multi, Missing, Missing, Missing, Missing, true);

		return ret;
	}



	inline void procedure(OPER& p)
	{
		// Strip leading underscore from C function
		if (p.val.str[1] == L'_') {
			p = OPER(p.val.str + 2, p.val.str[0] - 1);
		}
		// Prepend question mark for C++ name mangling.
		else if (p.val.str[1] != L'?') {
			p = OPER(L"?") & p;
		}
	}

	// Append "!0" to url if missing.
	inline void helpTopic(OPER& ht)
	{
		const auto help = view(ht);
		if (help.starts_with(L"http") && !help.ends_with(L"!0")) {
			ht &= OPER(L"!0");
		}
	}

	// Register a function or macro to be called by Excel.
	// https://learn.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1
	inline OPER XlfRegister(Args* pargs)
	{
		XLOPER12 res = { .xltype = xltypeNil };

		pargs->moduleText = Excel(xlGetName);
		procedure(pargs->procedure);
		helpTopic(pargs->helpTopic);

		constexpr size_t n = offsetof(Args, argumentHelp) / sizeof(OPER);
		const int count = n + size(pargs->argumentHelp);
		
		LPXLOPER12 as[255]; // array of pointers to arguments
		for (int i = 0; i < n; ++i) {
			as[i] = (LPXLOPER12)pargs + i;
		}
		//LPXLOPER12 pah = (LPXLOPER12)pargs->argumentHelp.val.array.lparray;
		LPXLOPER12 pah = Multi(pargs->argumentHelp);
		for (int i = 0; n + i < count; ++i) {
			as[n + i] = pah + i;
		}
		// https://docs.microsoft.com/en-us/office/client-developer/excel/known-issues-in-excel-xll-development#argument-description-string-truncation-in-the-function-wizard
		as[count] = const_cast<LPXLOPER12>(&Empty);

		const int ret = ::Excel12v(xlfRegister, &res, count, &as[0]);

		ensure_ret(ret); // call to Excel12v succeeded
		ensure_err(res); // call to xlfRegister succeeded
		ensure_message(type(res) == xltypeNum, "return type of xlfRegister must be the numeric RegisterId");

		return res;
	}
	
	// Really unregister a function.
	// https://learn.microsoft.com/en-us/office/client-developer/excel/xlfunregister-form-1
	// https://docs.microsoft.com/en-us/office/client-developer/excel/known-issues-in-excel-xll-development#unregistering-xll-commands-and-functions
	// https://stackoverflow.com/questions/15343282/how-to-remove-an-excel-udf-programmatically
	inline bool XlfUnregister(const OPER& procedure)
	{
		OPER regid = Excel(xlfEvaluate, procedure);
		if (type(regid) != xltypeNum) {
			OPER err(L"XlfUnregister: procedure not registered: ");
			XLL_WARNING(view(err & procedure));
			
			return false;
		}
		Excel(xlfSetName, procedure);
		
		regid = Excel(xlfRegister, Excel(xlGetName),
			OPER("xlAutoRemove"), OPER(XLL_SHORT), procedure, Missing, OPER(2));
		Excel(xlfSetName, procedure);

		return Excel(xlfUnregister, regid) == true;
	}

} // namespace xll
