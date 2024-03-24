// args.h - Arguments for Excel function and macro registration.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
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
			: type(type), name(name), help(help)
		{ }
	};

	// Arguments for xlfRegister.
	// https://learn.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1
	constexpr int max_help = 20;
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
		OPER argumentHelp[max_help]; // individual argument help
		/*
		Args(const Args&) = default;
		Args& operator=(const Args&) = default;
		~Args()
		{ }
		*/
	

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

		// Number of initial non Nil arguments
		int count() const
		{
			int i = 0;

			if (is_function()) {
				while (i < max_help && argumentHelp[i] != OPER{}) {
					++i;
				}
			}

			return i;
		}

		// Fix up arguments for call to xlfRegister.
		// Return count to be used in call to xlfRegister.
		int prepare()
		{
			int n = count();

			// idempotent
			if (moduleText == Nil) {
				moduleText = Excel(xlGetName);

				ensure(type(procedure) == xltypeStr);
				ensure(procedure.val.str[0] > 0);
				// C functions start with _
				if (procedure.val.str[1] == L'_') {
					procedure = OPER(procedure.val.str + 2, procedure.val.str[0] - 1);
				}
				// C++ name mangling.
				else if (procedure.val.str[1] != L'?') {
					procedure = OPER(L"?") & procedure;
				}

				if (is_function()) {
					if (helpTopic == Nil || helpTopic == L"") {
						helpTopic = OPER(L"https://google.com/search?q="); // github???
						helpTopic &= procedure;
					}

					const auto help = view(helpTopic);
					if (help.starts_with(L"http") && !help.ends_with(L"!0")) {
						helpTopic &= OPER(L"!0");
					}

					// Unpack typeText, argumentText, argumentHelp
					OPER comma(L"");
					for (int i = 0; i < n; ++i) {
						typeText = typeText & argumentHelp[i][Arg::Type::typeText];
						argumentText &= comma & argumentHelp[i][Arg::Type::argumentText];
						argumentHelp[i] = argumentHelp[i][Arg::Type::argumentHelp];
						comma = OPER(L", ");
					}
					// https://docs.microsoft.com/en-us/office/client-developer/excel/known-issues-in-excel-xll-development#argument-description-string-truncation-in-the-function-wizard
					//argumentHelp[n] = OPER(L"");
				}
			}

			constexpr size_t off = offsetof(Args, Args::argumentHelp) / sizeof(OPER);

			return static_cast<int>(off + n);
		}
	};

	struct Macro : public Args {
		template<class T> requires xll::is_char<T>::value
			Macro(const T* procedure, const T* functionText, const T* shortcut = nullptr)
			: Args{ .procedure = OPER(procedure),
					.functionText = OPER(functionText),
					.macroType = OPER(2),
					.shortcutText = shortcut ? OPER(shortcut) : OPER{} }
		{ }
	};

	struct Function : public Args {
		template<class T> requires is_char<T>::value
		Function(const wchar_t* type, const T* procedure, const T* functionText)
			: Args{ .procedure = OPER(procedure),
					.typeText = OPER(type),
					.functionText = OPER(functionText),
					.macroType = OPER(1) }
		{ }
		Function& Arguments(const std::initializer_list<Arg>& args)
		{
			OPER* ah = &argumentHelp[0];
			for (const auto& arg : args) {
				*ah++ = OPER({ arg.type, arg.name, arg.help });
			}

			return *this;
		}
		template<class T> requires is_char<T>::value
		Function& Category(const T* category_)
		{
			category = category_;

			return *this;
		}
		template<class T> requires xll::is_char<T>::value
		Function& FunctionHelp(const T* functionHelp_)
		{
			functionHelp = functionHelp_;

			return *this;
		}
		template<class T> requires xll::is_char<T>::value
		Function& HelpTopic(const T* helpTopic_)
		{
			helpTopic = helpTopic_;

			return *this;
		}
		Function& Uncalced()
		{
			typeText &= OPER(XLL_UNCALCED);

			return *this;
		}
		Function& Volatile()
		{
			typeText &= OPER(XLL_VOLATILE);

			return *this;
		}
		Function& ThreadSafe()
		{
			typeText &= OPER(XLL_THREAD_SAFE);

			return *this;
		}
		Function& Asynchronous()
		{
			typeText &= OPER(XLL_ASYNCHRONOUS);

			return *this;
		}
		Function& Hide()
		{
			macroType = OPER(0);

			return *this;
		}
	};

} // namespace xll