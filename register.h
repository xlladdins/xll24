// register.h - Excel function and macro registration.
#pragma once
#include "excel.h"

namespace xll {

	struct Arg {
		OPER type, name, text;
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
		OPER argumentHelp1;
		OPER argumentHelp2;
		OPER argumentHelp3;
		OPER argumentHelp4;
		OPER argumentHelp5;
		OPER argumentHelp6;
		OPER argumentHelp7;
		OPER argumentHelp8;
		OPER argumentHelp9;
		OPER argumentHelp10;
		OPER argumentHelp11;
		OPER argumentHelp12;
		OPER argumentHelp13;
		OPER argumentHelp14;
		OPER argumentHelp15;
		OPER argumentHelp16;
		OPER argumentHelp17;
		OPER argumentHelp18;
		OPER argumentHelp19;
		OPER argumentHelp20;
		OPER argumentHelp21;
		OPER argumentHelp22;
		int nargs = 0;
	};
	struct Macro : public Args {
		Macro(const XCHAR* procedure, const XCHAR* functionText, const XCHAR* shortcut = 0)
			: Args{ .procedure = OPER(procedure), .functionText = OPER(functionText), .macroType = OPER(2), .shortcutText = OPER(shortcut) }
		{ }
	};
	struct Function : public Args {
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

	inline OPER Register(Args args)
	{
		OPER res;
		LPXLOPER12 as[32];

		args.moduleText = Excel(xlGetName);
		ensure(type(args.procedure) == xltypeStr)
		if (args.procedure.val.str[1] != L'?') {
			args.procedure = Excel(xlfConcatenate, args.moduleText, OPER(L"?"), args.procedure);
		}
		auto arg = &args;

		for (int i = 0; i < 32; ++i) {
			as[i] = (LPXLOPER12)(arg + i);
		}

		ensure(xlretSuccess == Excel12(xlfRegister, &res, 30, &as[0]));

		return res;
	}

} // namespace xll
