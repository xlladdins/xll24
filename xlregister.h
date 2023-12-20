// xlregister.h - Excel registration functions
#pragma once
#include "xloper.h"

namespace xll {
	/*
	struct Function {
		OPER arg[32];
	};
	struct Macro {
	};
	*/

	inline OPER Register(const XLOPER12& ai)
	{
		OPER res;
		LPXLOPER12 as[32];

		for (int i = 0; i < 32; ++i) {
			as[i] = &ai[i];
		}

		ensure(xlretSuccess == Excel12(xlfRegister, &res, 32, &as[0]));

		return res;
	}

} // namespace xll
