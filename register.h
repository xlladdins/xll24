// register.h - Excel function and macro registration.
#pragma once
#include "excel.h"

namespace xll {

	// Return value of xlAddInManagerInfo.
	inline LPXLOPER12 AddInManagerInfo(const XLOPER12& info = Nil)
	{
		static OPER x;

		if (info.xltype == 2) { // xltypeStr
			x = info;
		}

		return &x;
	}

	struct Args {
		OPER moduleText;
		OPER procedure;
		OPER typeText;
		OPER functionText;
		OPER argumentText;
		OPER macroType;
		OPER category;
		OPER shortcutText;
		OPER helpTopic;
		OPER functionHelp;
		OPER argumentHelp1;
		OPER argumentHelp2;
		OPER argumentHelp3;
		OPER argumentHelp4;
		OPER argumentHelp5;
		OPER argumentHelp6;
		OPER argumentHelp7;
		OPER argumentHelp8;
		OPER argumentHelp9;
		OPER argumentHelp10;
		OPER argumentHelp11;
		OPER argumentHelp12;
		OPER argumentHelp13;
		OPER argumentHelp14;
		OPER argumentHelp15;
		OPER argumentHelp16;
		OPER argumentHelp17;
		OPER argumentHelp18;
		OPER argumentHelp19;
		OPER argumentHelp20;
		OPER argumentHelp21;
		OPER argumentHelp22;
		int nargs = 0;
	};

	inline OPER XlfRegister(Args args)
	{
		OPER res;
		LPXLOPER12 as[32];

		args.moduleText = Excel(xlGetName);
		ensure(type(args.procedure) == xltypeStr);
		if (args.procedure.val.str[1] != L'?') {
			args.procedure = OPER(L"?") & args.procedure;
		}

		auto arg = &args.moduleText;
		int i = 0;
		while (i < args.nargs && i < 31) {
			as[i] = (LPXLOPER12)(arg + i);
			++i;
		}
		// https://docs.microsoft.com/en-us/office/client-developer/excel/known-issues-in-excel-xll-development#argument-description-string-truncation-in-the-function-wizard
		OPER empty(L"");
		as[i] = &empty;

		int ret = Excel12v(xlfRegister, &res, args.nargs + 1, &as[0]);

		ensure_message(ret == xlretSuccess, xlret_description(ret));
		ensure_message(type(res) != xltypeErr, xlerr_description((xlerr)res.val.err));
		ensure_message(type(res) == xltypeNum, "return type of xlfRegister must be the numeric RegisterId");

		return res;
	}
	
	// Really unregister a function.
	// https://docs.microsoft.com/en-us/office/client-developer/excel/known-issues-in-excel-xll-development#unregistering-xll-commands-and-functions
	// https://stackoverflow.com/questions/15343282/how-to-remove-an-excel-udf-programmatically
	inline OPER Unregister(const OPER& key)
	{
		OPER regid = Excel(xlfEvaluate, key);
		if (type(regid) != xltypeNum) {
			return ErrValue;
		}
		Excel(xlfSetName, key);
		Excel(xlfUnregister, regid);

		regid = Excel(xlfRegister, Excel(xlGetName),
			OPER("xlAutoRemove"), OPER(XLL_SHORT), key, Missing, OPER(2));
		Excel(xlfSetName, key);

		return Excel(xlfUnregister, regid);
	}

} // namespace xll
