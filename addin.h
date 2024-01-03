// addin.h - create Excel add-ins to register
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <map>
#include "register.h"

namespace xll {

	struct Macro : public Args {
		template<class T> requires xll::is_char<T>::value
		Macro(const T* procedure, const T* functionText, const T* shortcut = nullptr)
			: Args{ .procedure = OPER(procedure),
					.functionText = OPER(functionText),
					.macroType = OPER(2.),
					.shortcutText = shortcut ? OPER(shortcut) : OPER{} }
		{ }
	};

	struct Function : public Args {
		template<class T> requires xll::is_char<T>::value
		Function(const T* type, const T* procedure, const T* functionText)
			: Args{ .procedure = OPER(procedure),
					.typeText = OPER(type),
					.functionText = OPER(functionText),
					.macroType = OPER(1.) }
		{ }
		Function& Arguments(const std::initializer_list<Arg>& args)
		{
			OPER* ah = &argumentHelp[0];
			for (const auto& arg : args) {
				*ah++ = OPER({ arg.type, arg.name, arg.help });
			}

			return *this;
		}
		template<class T> requires xll::is_char<T>::value
		Function& Category(const T* category_)
		{
			category = OPER(category_);

			return *this;
		}
		template<class T> requires xll::is_char<T>::value
		Function& FunctionHelp(const T* functionHelp_)
		{
			functionHelp[0] = OPER(functionHelp_);

			return *this;
		}
		template<class T> requires xll::is_char<T>::value
		Function& HelpTopic(const T* helpTopic_)
		{
			helpTopic = OPER(helpTopic_);

			return *this;
		}
		Function& Uncalced()
		{
			typeText = OPER(XLL_UNCALCED) & typeText;

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
	};

	struct AddIn {
		//static inline std::map<OPER, Args> addins;
		AddIn(const Args& args)
		{
			//if (addins.contains(args.functionText)) {
				//std::string msg = L"AddIn: " + view(args.functionText) + L" already registered";
				//XLL_WARNING(std::string("AddIn: ") + view(args.functionText) + " already registered");
			//}
			Auto<Register> reg([args]() {
				OPER regId = XlfRegister(args);

				return regId.xltype == xltypeNum; });
		}
		AddIn(const AddIn&) = delete;
		AddIn& operator=(const AddIn&) = delete;
		~AddIn()
		{
			//Auto<Remove> unreg([]() { return XlfUnregister(); };
		}
		// Auto<Unregister>
		// Auto<Close>
	};

} // namespace xll
