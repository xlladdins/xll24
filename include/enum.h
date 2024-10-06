// enum.h - Functions used for enumerations.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <expected>
#include "oper.h"
#include "handle.h"

namespace xll
{
	constexpr bool StartsWith(const XLOPER12& x, int c)
	{
		return isStr(x) && x.val.str[0] > 0 && x.val.str[1] == c;
	}
	constexpr bool EndsWith(const XLOPER12& x, int c)
	{
		return isStr(x) && x.val.str[0] > 0 && x.val.str[x.val.str[0]] == c;
	}
	constexpr bool isFormula(const XLOPER12& x)
	{
		return StartsWith(x, '=');
	}
	// "=...)"
	constexpr bool isFunction(const XLOPER12& x)
	{
		return isFormula(x) && EndsWith(x, ')');
	}	
	// "\..."
	// Functions returning a handle start with backslash by convention.
	constexpr bool isHandle(const XLOPER12& x)
	{
		return StartsWith(x, '\\');
	}
#ifdef _DEBUG
	static_assert(StartsWith(Str(L"\x03=A1"), '='));
	static_assert(EndsWith(Str(L"\x03=A1"), '1'));
	static_assert(isFormula(Str(L"\x03=A1")));
	static_assert(!isFunction(Str(L"\x03=A1")));
	static_assert(isHandle(Str(L"\x03\\A1")));
#endif // _DEBUG

	// String is name of a user defined function
	inline bool isUDF(const XLOPER12& x)
	{
		return isStr(x) && RegIds().contains(Num(Excel(xlfEvaluate, x)));
	}
	// UDF with no arguments
	inline bool isEnum(const XLOPER12& x)
	{
		const Args* pargs = AddIn::find(x);
		if (pargs && size(pargs->argumentType) == 0) {
			return true;
		}

		return false;
	}

	inline OPER Eval(const XLOPER12& x)
	{
		OPER o = x;

		if (isEnum(x)) {
			o = Excel(xlfEvaluate, OPER(L"=") & OPER(x) & OPER(L"()"));
		}
		else if (isFormula(x)) {
			o = Excel(xlfEvaluate, x);
		}

		return o;
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
