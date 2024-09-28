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
		OPER argumentHelp; // array of individual argument help
		// Auxiliary arrays.
		OPER argumentType; // array of individual argument types
		OPER argumentName; // array of individual argument names	
		OPER argumentInit; // array individual argument default values
		OPER documentation;

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

		Args& Documentation(std::string_view doc)
		{
			documentation = doc;

			return *this;
		}
		Args& Documentation(std::wstring_view doc)
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
			o.vstack(OPER({ OPER(L"argumentHelp"), argumentHelp }));
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
			OPER comma(L"");
			for (const auto& arg : args) {
				typeText &= arg.type;
				argumentText &= (comma & arg.name);

				argumentHelp.append(arg.help);
				argumentType.append(arg.type);
				argumentName.append(arg.name);
				argumentInit.append(arg.init);

				comma = OPER(L", ");
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