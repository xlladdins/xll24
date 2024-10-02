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
#define XLL_REGISTER_ARGS(X) \
X(moduleText, xltypeStr, "The name of the DLL that contains the function.") \
X(procedure, xltypeStr, "The name of the function to call as it appears in the DLL code.") \
X(typeText, xltypeStr, "The type of the function.") \
X(functionText, xltypeStr, "The name of the function as it appears in the Excel function wizard.") \
X(argumentText, xltypeStr, "The argument text description for the function.") \
X(macroType, xltypeInt, "The type of the function: 0 for hidden, 1 for function, 2 for macro.") \
X(category, xltypeStr|xltypeNum, "The category of the function in the Function Wizard.") \
X(shortcutText, xltypeStr, "A one-character, case-sensitive string that specifies the control key assigned to this command.") \
X(helpTopic, xltypeStr, "The URL of the Help topic for the function.") \
X(functionHelp, xltypeStr, "The help text for the function.") \
X(argumentHelp, xltypeMulti, "The help text for each argument.") \
X(argumentType, xltypeMulti, "The type of each argument.") \
X(argumentName, xltypeMulti, "The name of each argument.") \
X(argumentInit, xltypeMulti, "The default value of each argument.") \
X(documentation, xltypeStr, "The documentation for the function.") \

	enum class args {
#define XLL_REGISTER_ARG(name, type, help) name,
		XLL_REGISTER_ARGS(XLL_REGISTER_ARG)
#undef XLL_REGISTER_ARG
	};

	constexpr const char* arg_name(args arg)
	{
		switch (arg) {
#define XLL_REGISTER_ARG(name, type, help) case args::name: return #name;
		XLL_REGISTER_ARGS(XLL_REGISTER_ARG)
#undef XLL_REGISTER_ARG
		default:
			return nullptr;
		}
	}

	struct Args {
#define XLL_REGISTER_ARG(name, type, help) OPER name;
		XLL_REGISTER_ARGS(XLL_REGISTER_ARG)
#undef XLL_REGISTER_ARG

		constexpr size_t size() const
		{
			return sizeof(Args) / sizeof(OPER);
		}
		// Assumes lifetime of Args
		constexpr const XLOPER12 asMulti() const
		{
			return Multi((XLOPER12*)&moduleText, static_cast<int>(size()), 1);
		}

		constexpr OPER& operator[](args arg)
		{
			return *(&moduleText + static_cast<int>(arg));
		}

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