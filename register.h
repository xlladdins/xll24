// register.h - Excel function and macro registration.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
// https://learn.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1
// Excel12(xlfRegister, LPXLOPER12 pxRes, int iCount,
//	LPXLOPER12 pxModuleText, LPXLOPER12 pxProcedure,
//	LPXLOPER12 pxTypeText, LPXLOPER12 pxFunctionText,
//	LPXLOPER12 pxArgumentText, LPXLOPER12 pxMacroType,
//	LPXLOPER12 pxCategory, LPXLOPER12 pxShortcutText,
//	LPXLOPER12 pxHelpTopic, LPXLOPER12 pxFunctionHelp,
//	LPXLOPER12 pxArgumentHelp1, LPXLOPER12 pxArgumentHelp2,
//	...);

#pragma once
#include "excel.h"

namespace xll {

	// Individual argument for an add-in function.
	struct Arg {
		enum Type {
			typeText, argumentText, argumentHelp
		};
		OPER type, name, help;
		template<class T> 
			requires xll::is_char<T>::value
		Arg(const wchar_t* type, const T* name, const T* help)
			: type(OPER(type)), name(OPER(name)), help(OPER(help))
		{ }
	};

	// Arguments for xlfRegister.
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
		OPER argumentHelp[20];

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

		// Number of non Nil arguments
		int count() const
		{
			int i = 0;

			if (is_function()) {
				while (i < 20 && argumentHelp[i] != Nil) {
					++i;
				}
			}

			return i;
		}
		
		// Fix up arguments for call to xlfRegister.
		// Return count to be used in call to xlfRegister.
		int prepare() 
		{
			moduleText = Excel(xlGetName);

			// C++ name mangling.
			ensure(type(procedure) == xltypeStr);
			ensure(procedure.val.str[0] > 0);
			if (procedure.val.str[1] != L'?') {
				procedure = OPER(L"?") & procedure;
			}

			int n = count();
			if (is_function()) {
				if (helpTopic == Nil) {
					helpTopic = OPER(L"https://google.com/search?q="); // github???
					helpTopic &= procedure;
				}

				const auto help = view(helpTopic);
				if (help.starts_with(L"http") && !help.ends_with(L"!0")) {
					helpTopic &= OPER(L"!0");
				}

				// typeText, argumentText, functionHelp
				OPER comma = OPER(L"");
				for (int i = 0; i < n; ++i) {
					typeText = typeText & argumentHelp[i][Arg::Type::typeText];
					argumentText &= comma & argumentHelp[i][Arg::Type::argumentText];
					argumentHelp[i] = argumentHelp[i][Arg::Type::argumentHelp];
					comma = OPER(L", ");
				}
				// https://docs.microsoft.com/en-us/office/client-developer/excel/known-issues-in-excel-xll-development#argument-description-string-truncation-in-the-function-wizard
				argumentHelp[n] = OPER(L"");
			}

			return static_cast<int>(offsetof(Args, argumentHelp) / sizeof(OPER) + n + 1);
		}
	};

	// Register a function or macro to be called by Excel.
	// https://learn.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1
	inline OPER XlfRegister(const Args& args_)
	{
		OPER res;

		Args args(args_);
		constexpr size_t n = sizeof(Args)/ sizeof(OPER);
		LPXLOPER12 as[n];
		const int count = args.prepare();
		for (int i = 0; i < n && i < count; ++i) {
			as[i] = (LPXLOPER12)&args + i;
		}

		const int ret = ::Excel12v(xlfRegister, &res, count, &as[0]);

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
		
		regid = Excel(xlfRegister, Excel(xlGetName),
			OPER("xlAutoRemove"), OPER(XLL_SHORT), key, Missing, OPER(2));
		Excel(xlfSetName, key);

		return Excel(xlfUnregister, regid) == true;
	}

} // namespace xll
