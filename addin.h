// addin.h - create Excel add-ins to register
#pragma once
#include "register.h"

namespace xll {

	struct Macro : public Args {
		Macro(const XCHAR* procedure, const XCHAR* functionText, const XCHAR* shortcut = nullptr)
			: Args{ .procedure = OPER(procedure),
					.functionText = OPER(functionText),
					.macroType = OPER(2.),
					.shortcutText = shortcut ? OPER(shortcut) : OPER{},
					.nargs = 10 }
			{ }
	};

	struct Arg {
		OPER type, name, text, init;
		Arg(const XCHAR* type, const char* name, const char* text, const OPER& init = Missing)
			: type(OPER(type)), name(OPER(name)), text(OPER(text)), init(init)
		{ }
	};

	struct Function : public Args {
		Function(const XCHAR* type, const XCHAR* procedure, const XCHAR* functionText)
			: Args{ .procedure = OPER(procedure),
			        .typeText = OPER(type),
					.functionText = OPER(functionText),
					.macroType = OPER(1),
			        .nargs = 10}
		{ }
		Function& Arguments(const std::initializer_list<Arg>& args)
		{
			OPER* fh = &functionHelp;
			OPER comma = OPER(L"");

			for (auto& arg : args) {
				ensure(nargs < 22);
				typeText &= arg.type;
				functionText = comma & functionText;
				*fh++ = arg.text;
				comma = OPER(L", ");
				++nargs;
			}

			return *this;
		}
		Function& Uncalced()
		{
			typeText = OPER(XLL_UNCALCED) & typeText;

			return *this;
		}
		Function& Volatile()
		{
			typeText = OPER(XLL_VOLATILE) & typeText;

			return *this;
		}
		Function& ThreadSafe()
		{
			typeText = OPER(XLL_THREAD_SAFE) & typeText;

			return *this;
		}
		Function& Asynchronous()
		{
			typeText = OPER(XLL_ASYNCHRONOUS) & typeText;

			return *this;
		}
	};

	struct AddIn {
		AddIn(const Args& args)
		{
			Auto<Open> reg([args]() { return XlfRegister(args).xltype == xltypeNum; });
		}
		AddIn(const AddIn&) = delete;
		AddIn& operator=(const AddIn&) = delete;
		~AddIn()
		{
			//Auto<Unregister> unregister;
		}
		// Auto<Unregister>
		// Auto<Close>
	};

} // namespace xll
