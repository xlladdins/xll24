// addin.h - create Excel add-ins to register
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <map>
#include "handle.h"
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
		Function(const wchar_t* type, const T* procedure, const T* functionText)
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
	};

	struct AddIn {
		static inline std::map<OPER, Args> addins;
		AddIn(const Args& args)
		{
			if (const auto e = addins.find(args.functionText); e != addins.end()) {
				const auto err = OPER(L"AddIn: ") 
					& args.functionText & OPER(L" already registered as: ") & e->second.procedure;
				XLL_WARNING(view(err));
			}
			else {
				addins[args.functionText] = args;
			}
			Auto<Register> reg([&args]() {
				OPER regId = XlfRegister(args);

				return regId.xltype == xltypeNum; 
			});
			Auto<Unregister> unreg([]() {
				try {
					for (const auto& [text, args] : addins) {
						if (!XlfUnregister(text)) {
							const auto err = OPER(L"AddIn: failed to unregister: ") & text;
							XLL_WARNING(view(err));

							return FALSE;
						}
					}
				}
				catch (const std::exception& ex) {
					XLL_ERROR(ex.what());

					return FALSE;
				}
				catch (...) {
					XLL_ERROR("AddIn::Auto<Close>: unknown exception");

					return FALSE;
				}

				return TRUE;
			});
		}
		AddIn(const AddIn&) = delete;
		AddIn& operator=(const AddIn&) = delete;
		~AddIn()
		{ }
	};

} // namespace xll
