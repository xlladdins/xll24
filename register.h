// xlregister.h - Excel registration functions
#pragma once
#include "excel.h"

namespace xll {
	/*
	struct Function {
		OPER arg[32];
	};
	struct Macro {
	};
	*/
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
	};
	struct Macro : public Args {
		Macro(const XCHAR* procedure, const XCHAR* functionText, const XCHAR* shortcut)
			: Args{ .procedure = OPER(procedure), .functionText = OPER(functionText), .macroType = OPER(2), .shortcutText = OPER(shortcut) }
		{ }
	};
	struct Function : public Args {
	};

	inline OPER Register(Args args)
	{
		OPER res;
		LPXLOPER12 as[32];
		args.moduleText = Excel(xlGetName);
		auto arg = &args.moduleText;

		for (int i = 0; i < 32; ++i) {
			as[i] = (LPXLOPER12)(arg + i);
		}

		ensure(xlretSuccess == Excel12(xlfRegister, &res, 30, &as[0]));

		return res;
	}

} // namespace xll
