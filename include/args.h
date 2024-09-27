// args.h - Arguments for Excel function and macro registration.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include "excel.h"

namespace xll {

	// Individual argument for an add-in function.
	struct Arg {
		OPER type, name, help, init;

		template<class T>
			requires xll::is_char<T>::value
		Arg(const wchar_t* type, const T* name, const T* help)
			: type(type), name(name), help(help), init(OPER())
		{ }
		template<class T, class U>
			requires xll::is_char<T>::value
		Arg(const wchar_t* type, const T* name, const T* help, U init)
			: type(type), name(name), help(help), init(init)
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
		OPER argumentHelp[max_help]; // array of individual argument help
		// Auxiliary arrays.
		OPER argumentType; // array of individual argument types
		OPER argumentName; // array of individual argument names	
		OPER argumentInit; // array individual argument default values
		std::string documentation;

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
				while (i < max_help && !isNil(argumentHelp[i])) {
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
				moduleText = Excel(xlGetName); // Can only be called after xlAutoOpen.

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

					// Unpack typeText, argumentText
					OPER comma(L"");
					for (int i = 0; i < n; ++i) {
						typeText &= argumentType[i];
						argumentText &= (comma & argumentName[i]);
						comma = OPER(L", ");
					}

					// https://docs.microsoft.com/en-us/office/client-developer/excel/known-issues-in-excel-xll-development#argument-description-string-truncation-in-the-function-wizard
					//argumentHelp[n] = OPER(L"");
				}
			}

			size_t off = offsetof(Args, Args::argumentHelp) / sizeof(OPER);

			return static_cast<int>(off + n);
		}
		Args& Documentation(std::string_view doc)
		{
			documentation = doc;

			return *this;
		}

		// key-value pairs
		OPER Info() const
		{
			OPER o({
				OPER(L"moduleText"), moduleText,
				OPER(L"procedure"), procedure,
				OPER(L"typeText"), typeText,
				OPER(L"functionText"), functionText,
				OPER(L"argumentText"), argumentText,
				OPER(L"macroType"), macroType,
				OPER(L"category"), category,
				OPER(L"shortcutText"), shortcutText,
				OPER(L"helpTopic"), helpTopic,
				OPER(L"functionHelp"), functionHelp 
			});
			o.reshape(size(o) / 2, 2);
			OPER ah;
			for (int i = 0; i < count(); ++i) {
				ah.append(argumentHelp[i]);
			}
			o.vstack(OPER({ OPER(L"argumentHelp"), ah}));
			o.vstack(OPER({ OPER(L"argumentType"), argumentType }));
			o.vstack(OPER({ OPER(L"argumentName"), argumentName }));
			o.vstack(OPER({ OPER(L"argumentInit"), argumentInit }));
			o.vstack(OPER({ OPER(L"documentation"), OPER(documentation) }));

			return o;
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
		Macro(const Macro&) = default;
		Macro& operator=(const Macro&) = default;
		~Macro()
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
		Function(const Function&) = default;
		Function& operator=(const Function&) = default;
		~Function()
		{ }
		Function& Arguments(const std::initializer_list<Arg>& args)
		{
			int i = 0;
			for (const auto& arg : args) {
				ensure(i < max_help);
				argumentHelp[i++] = arg.help;
				argumentType.append(arg.type);
				argumentName.append(arg.name);
				argumentInit.append(arg.init);
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