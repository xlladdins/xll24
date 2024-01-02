// register.h - Excel function and macro registration.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include "excel.h"

namespace xll {

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
		
		int prepare() // for call to xlfRegister
		{

			moduleText = Excel(xlGetName);

			// C++ name mangling.
			ensure(type(procedure) == xltypeStr);
			if (procedure.val.str[1] != L'?') {
				procedure = OPER(L"?") & procedure;
			}

			//!!! search github ???
			if (helpTopic == Nil) {
				helpTopic = OPER(L"https://google.com/search?q=");
				helpTopic &= procedure;
			}
			if (view(helpTopic).starts_with(L"http")) {
				helpTopic &= OPER(L"!0");
			}

			int i = 1;
			if (functionHelp[0] == Nil) { // macro
				//functionHelp[/*++*/0] = OPER(L"");
			}
			else { // function
				// typeText, argumentText, functionHelp
				OPER comma = OPER(L"");
				while (functionHelp[i].xltype != xltypeNil) {
					typeText = typeText & functionHelp[i][Arg::Type::typeText];
					argumentText &= comma & functionHelp[i][Arg::Type::argumentText];
					functionHelp[i] = functionHelp[i][Arg::Type::functionHelp];
					comma = OPER(L", ");
					++i;
				}
				// https://docs.microsoft.com/en-us/office/client-developer/excel/known-issues-in-excel-xll-development#argument-description-string-truncation-in-the-function-wizard
				functionHelp[i] = OPER(L"");
			}

			return static_cast<int>(offsetof(Args, functionHelp)/sizeof(OPER) + i + 1);
		}
	};

	// https://learn.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1
	inline OPER XlfRegister(Args args)
	{
		OPER res;

		LPXLOPER12 as[32];
		int count = args.prepare();
		for (int i = 0; i < count; ++i) {
			as[i] = (LPXLOPER12) &args + i;
		}

		int ret = ::Excel12v(xlfRegister, &res, count, &as[0]);

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
