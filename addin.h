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
					.nargs = 8 }
		{ }
	};

	struct Arg {
		OPER type, name, text, init;
		Arg(const XCHAR* type, const char* name, const char* text, const OPER& init = Nil())
			: type(OPER(type)), name(OPER(name)), text(OPER(text)), init(init)
		{ }
	};

	struct Function : public Args {
		Function(const XCHAR* procedure, const XCHAR* functionText)
			: Args{ .procedure = OPER(procedure),
					.functionText = OPER(functionText),
					.macroType = OPER(1) }
		{ }
		Function& Arguments(const std::initializer_list<Arg>& args)
		{
			int i = 0;
			OPER* fh = &functionHelp;
			OPER comma = OPER(L"");
			for (auto& arg : args) {
				ensure(nargs < 22);
				typeText = typeText & arg.type;
				functionText = comma & functionText;
				fh[i] = arg.text;
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

	inline void AddIn(const Args& args)
	{
		Auto<Register> reg([args]() { return XlfRegister(args).xltype == xltypeNum; });
		// Auto<Unregister>
		// Auto<Close>
	}


} // namespace xll
