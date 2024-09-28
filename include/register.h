// register.h - Excel function and macro registration.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include "args.h"

namespace xll {

	// Register a function or macro to be called by Excel.
	// https://learn.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1
	inline OPER XlfRegister(Args* pargs)
	{
		XLOPER12 res = { .xltype = xltypeNil };

		pargs->moduleText = Excel(xlGetName);
		if (pargs->procedure.val.str[1] == L'_') {
			pargs->procedure 
				= OPER(pargs->procedure.val.str + 2, pargs->procedure.val.str[0] - 1);
		}
		// C++ name mangling.
		else if (pargs->procedure.val.str[1] != L'?') {
			pargs->procedure = OPER(L"?") & pargs->procedure;
		}

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
