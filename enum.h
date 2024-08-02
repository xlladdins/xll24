// enum.h - define XLL_CONST functions used for enumerations
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <expected>
#include "oper.h"
#include "handle.h"

namespace xll
{
	constexpr bool StartsWith(const XLOPER12& x, wchar_t c)
	{
		return isStr(x) && x.val.str[0] > 0 && x.val.str[1] == c;
	}
	constexpr bool EndsWith(const XLOPER12& x, wchar_t c)
	{
		return isStr(x) && x.val.str[0] > 0 && x.val.str[x.val.str[0]] == c;
	}
	// "=...)"
	inline bool isFormula(const XLOPER12& x)
	{
		return StartsWith(x, L'=') && EndsWith(x, L')');	
	}	

	// "\..."
	// Functions returning a handle start with backslash by convention.
	inline bool isHandle(const XLOPER12& x)
	{
		return StartsWith(x, L'\\');
	}

	// String is name of a user defined function
	inline bool isUDF(const XLOPER12& x)
	{
		return isStr(x) && AddIns().contains(x);
	}

	inline OPER Eval(const XLOPER12& x)
	{
		return Excel(xlfEvaluate, isUDF(x) ? OPER(L"=") & x & OPER(L"()") : x);
	}

	// Return asNum(o) as T. If o is a string, evaluate it first.
	template<class T>
	inline std::expected<T, OPER> EnumVal(const XLOPER12& o, T init)
	{
		if (isTrue(o)) {
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
	// Convert to safe pointer.
	template<class T>
	inline std::expected<T*, OPER> EnumPtr(const XLOPER12& o, T* init)
	{
		if (isTrue(o)) {
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
