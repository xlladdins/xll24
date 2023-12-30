// register.h - Excel function and macro registration.
#pragma once
#include "excel.h"

namespace xll {

	// Return value of xlAddInManagerInfo.
	inline LPXLOPER12 AddInManagerInfo(const XLOPER12& info = Nil)
	{
		static OPER x;

		if (info.xltype == xltypeStr) {
			x = info;
		}

		return &x;
	}

	struct Arg {
		enum Type {
			typeText, argumentText, functionHelp
		};
		OPER type, name, text;
		template<class T> requires xll::is_char<T>::value
			Arg(const T* type, const T* name, const T* text)
			: type(OPER(type)), name(OPER(name)), text(OPER(text))
		{ }
	};

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
		OPER functionHelp[23];
	};

	// https://learn.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1
	inline OPER XlfRegister(Args args)
	{
		OPER res;

		// moduleText
		args.moduleText = Excel(xlGetName);

		// procedure
		ensure(type(args.procedure) == xltypeStr);
		if (args.procedure.val.str[1] != L'?') {
			args.procedure = OPER(L"?") & args.procedure;
		}

		// typeText, argumentText, functionHelp
		int i = 0;
		OPER comma = OPER(L"");
		while (args.functionHelp[i].xltype != xltypeNil) {
			args.typeText = args.typeText & args.functionHelp[i][Arg::Type::typeText];
			args.argumentText &= comma & args.functionHelp[i][Arg::Type::argumentText];
			args.functionHelp[i] = args.functionHelp[i][Arg::Type::functionHelp];
			comma = OPER(L", ");
			++i;
		}

		LPXLOPER12 as[32];
		XLOPER12* arg = &args.moduleText;
		std::transform(arg, arg + 9 + i, as, [](XLOPER12& o) { return &o; });

		// https://docs.microsoft.com/en-us/office/client-developer/excel/known-issues-in-excel-xll-development#argument-description-string-truncation-in-the-function-wizard
		OPER empty(L"");
		as[9 + i + 1] = &empty;

		int ret = ::Excel12v(xlfRegister, &res, 9 + i + 1, &as[0]);

		ensure_ret(ret);
		ensure_err(res);
		ensure_message(type(res) == xltypeNum, "return type of xlfRegister must be the numeric RegisterId");

		return res;
	}
	
	// Really unregister a function.
	// https://docs.microsoft.com/en-us/office/client-developer/excel/known-issues-in-excel-xll-development#unregistering-xll-commands-and-functions
	// https://stackoverflow.com/questions/15343282/how-to-remove-an-excel-udf-programmatically
	inline OPER XlfUnregister(const OPER& key)
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
