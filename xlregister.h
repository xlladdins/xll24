// xlregister.h - Excel registration functions
#pragma once
#include "xlexcel.h"

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
	};
	struct Macro : public Args {
		Macro(const OPER& procedure, const OPER& functionText, const OPER& shortcut)
			: Args{ .procedure = procedure, .functionText = functionText, .macroType = OPER(2), .shortcutText = shortcut }
		{ }
	};

	inline OPER Register(Args args)
	{
		OPER res;
		LPXLOPER12 as[32];
		ensure(xlretSuccess == Excel12v(xlGetName, &args.moduleText, 0, 0));
		auto arg = &args.moduleText;

		for (int i = 0; i < 32; ++i) {
			as[i] = (LPXLOPER12)(arg + i);
		}

		ensure(xlretSuccess == Excel12(xlfRegister, &res, 32, &as[0]));

		return res;
	}

} // namespace xll
