// xlregister.h - Excel registration functions
#pragma once
#include "xloper.h"

namespace xll {
	/*
	inline OPER12 Function(const OPER12& name, const OPER12& args, const OPER12& macro)
	{
		OPER12 x;

		x.xltype = xltypeStr;
		x.val.str = name.val.str;
		Excel12(xlfRegister, 0, 3, &x, &args, &macro);

		return x;
	}
	*/

	inline OPER12 Register(const XLOPER12& ai)
	{
		OPER12 res;
		XLOPER12 as[32];

		for (int i = 0; i < 32; ++i) {
			as[i] = &ai[i];
		}

		ensure(xlretSuccess == Excel12(xlfRegister, &res, 32, &as[0]));

		return res;
	}

} // namespace xll
