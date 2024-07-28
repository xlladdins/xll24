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
		if (isMissing(e) || isNil(e) || (isNum(e) && Num(e) == 0)) {
			return init;
		}

		ensure(isNum(e) || isStr(e));

		HANDLEX h = INVALID_HANDLEX;
		if (isNum(e)) {
			T t = safe_value<T>(Num(e));
			if (t == 0) {
				init = (T)Num(e);
			}
		}	
		else if (isStr(e)) {
			OPER o = Excel(xlfEvaluate, OPER(L"=") & e & OPER(L"()"));
		}

		return init;
	}
}
