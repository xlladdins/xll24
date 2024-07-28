// enum.h - define XLL_CONST functions used for enumerations
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include "oper.h"
#include "handle.h"

namespace xll
{
	// Convert OPER to type defined in XLL_CONST, or default value.
	template<class T>
	inline T Enum(const OPER& e, T init) 
	{
		if (isNum(e) && Num(e) != 0) {
			init = safe_value<T>(Num(e));
			ensure(init);
		}	
		else if (isStr(e)) {
			OPER o = Excel(xlfEvaluate, OPER(L"=") & e & OPER(L"()"));
			ensure(isNum(o));
			return Enum(o, init);
		}

		return init;
	}
}
