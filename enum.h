// enum.h - define XLL_CONST functions used for enumerations
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <expected>
#include "oper.h"
#include "handle.h"

namespace xll
{
	// Evaluate an enumerated constant.
	inline OPER Eval(const OPER& o)
	{
		return Excel(xlfEvaluate, OPER(L"=") & o & OPER(L"()"));
	}

	// Return asNum(o) as T. If o is a string, evaluate it first.
	template<class T>
	inline std::expected<T, OPER> EnumVal(const OPER& o, T init)
	{
		if (o) {
			double x;
			if (isStr(o)) {
				OPER e = Eval(o);
				if (!isNum(e)) {
					return std::unexpected<OPER>(e);
				}
				x = asNum(e);
			}
			else {
				x = asNum(o);
			}
			if (std::isnan(x)) {
				return std::unexpected<OPER>(ErrValue);
			}
			init = static_cast<T>(x);

		}

		return init;
	}
	template<class T>
	inline std::expected<T*, OPER> EnumPtr(const OPER& o, T* init)
	{
		if (o) {
			HANDLEX h = INVALID_HANDLEX;
			if (isStr(o)) {
				OPER e = Eval(o);
				if (!isNum(e)) {
					return std::unexpected(e);
				}
				h = asNum(e);
			}
			else {
				h = asNum(o);
			}
			if (std::isnan(h)) {
				return std::unexpected(ErrValue);
			}

			init = safe_pointer<T>(h);
		}

		return init;
	}

}
