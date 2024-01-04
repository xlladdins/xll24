// register.h - Excel function and macro registration.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include "handle.h"

namespace xll {


	struct Arg {
		enum Type {
			typeText, argumentText, argumentHelp
		};
		OPER type, name, help;
		template<class T> requires xll::is_char<T>::value
			Arg(const T* type, const T* name, const T* help)
			: type(OPER(type)), name(OPER(name)), help(OPER(help))
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
		OPER functionHelp;
		OPER argumentHelp[22];

		bool is_hidden() const
		{
			return macroType == 0;
		}
		bool is_function() const
		{
			return macroType == 0 || macroType == 1;
		}
		bool is_macro() const
		{
			return macroType == 2;
		}
		
		int prepare() // for call to xlfRegister
		{

			moduleText = Excel(xlGetName);

			// C++ name mangling.
			ensure(type(procedure) == xltypeStr);
			if (procedure.val.str[1] != L'?') {
				procedure = OPER(L"?") & procedure;
			}

			int i = 0;
			//!!! search github ???
			if (is_function()) {
				if (helpTopic == Nil) {
					helpTopic = OPER(L"https://google.com/search?q=");
					helpTopic &= procedure;
				}
				if (view(helpTopic).starts_with(L"http")) {
					helpTopic &= OPER(L"!0");
				}

				// typeText, argumentText, functionHelp
				OPER comma = OPER(L"");
				while (i < 22 && argumentHelp[i] != Nil) {
					typeText = typeText & argumentHelp[i][Arg::Type::typeText];
					argumentText &= comma & argumentHelp[i][Arg::Type::argumentText];
					argumentHelp[i] = argumentHelp[i][Arg::Type::argumentHelp];
					comma = OPER(L", ");
					++i;
				}
			}
			// https://docs.microsoft.com/en-us/office/client-developer/excel/known-issues-in-excel-xll-development#argument-description-string-truncation-in-the-function-wizard
			argumentHelp[i] = OPER(L"");

			return static_cast<int>(offsetof(Args, argumentHelp)/sizeof(OPER) + i + 1);
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
	// https://learn.microsoft.com/en-us/office/client-developer/excel/xlfunregister-form-1
	// https://docs.microsoft.com/en-us/office/client-developer/excel/known-issues-in-excel-xll-development#unregistering-xll-commands-and-functions
	// https://stackoverflow.com/questions/15343282/how-to-remove-an-excel-udf-programmatically
	inline bool XlfUnregister(const OPER& key)
	{
		OPER regid = Excel(xlfEvaluate, key);
		if (type(regid) != xltypeNum) {
			return false;
		}
		Excel(xlfSetName, key);
		
		return Excel(xlfUnregister, regid) == true;
		/* Is this really necessary???
		regid = Excel(xlfRegister, Excel(xlGetName),
			OPER("xlAutoRemove"), OPER(XLL_SHORT), key, Missing, OPER(2));
		Excel(xlfSetName, key);

		return Excel(xlfUnregister, regid) == true;
		*/
	}

} // namespace xll
