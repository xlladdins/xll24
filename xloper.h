// xloper.h - XLOPER12 helpers
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include "ref.h"

namespace xll {

	// return XLFree(x) in thread-safe functions
	// Freed by Excel when no longer needed.
	constexpr LPXLOPER12 XLFree(XLOPER12& x) noexcept
	{
		x.xltype |= xlbitXLFree;

		return &x;
	}
	// return DLLFree(x) in thread-safe functions
	// Excel calls xlAutoFree12 when no longer needed.
	constexpr LPXLOPER12 DLLFree(XLOPER12& x) noexcept
	{
		x.xltype |= xlbitDLLFree;

		return &x;
	}

	constexpr double as_num(const XLOPER12& x) noexcept
	{
		switch (type(x)) {
		case xltypeNum:
			return x.val.num;
		case xltypeBool:
			return x.val.xbool;
		case xltypeInt:
			return x.val.w;
		case xltypeMissing:
		case xltypeNil:
			return 0;
		}

		return std::numeric_limits<double>::quiet_NaN();
	}

	constexpr int rows(const XLOPER12& x) noexcept
	{
		switch (type(x)) {
		case xltypeMulti:
			return x.val.array.rows;
		case xltypeSRef:
			return rows(SRef(x));
		case xltypeMissing:
		case xltypeNil:
			return 0;
		}

		return 1;
	}
	constexpr int columns(const XLOPER12& x) noexcept
	{
		switch (type(x)) {
		case xltypeMulti:
			return x.val.array.columns;
		case xltypeSRef:
			return columns(SRef(x));
		case xltypeMissing:
		case xltypeNil:
			return 0;
		}

		return 1;
	}
	constexpr int size(const XLOPER12& x) noexcept
	{
		return rows(x) * columns(x);
	}

	constexpr std::wstring_view view(const XLOPER12& x)
	{
		return std::wstring_view(Str(x), count(x));
	}

	constexpr auto span(const XLOPER12& x)
	{
		return std::span(Multi(x), count(x));
	}

	constexpr auto ref(const XLOPER12& x)
	{
		return std::span(Ref(x), count(x));
	}

	constexpr auto blob(const XLOPER12& x)
	{
		return std::span(BigData(x), count(x));
	}

	constexpr std::partial_ordering compare(const XLOPER12& x, const XLOPER12& y)
	{
		int xtype = type(x);
		int ytype = type(y);

		if (xtype != ytype) {
			return xtype <=> ytype;
		}

		switch (xtype) {
		case xltypeNum:
			return x.val.num <=> y.val.num;
		case xltypeStr:
			return std::lexicographical_compare_three_way(view(x).begin(), view(x).end(), view(y).begin(), view(y).end());
		case xltypeBool:
			return x.val.xbool <=> y.val.xbool;
		case xltypeErr:
			return x.val.err <=> y.val.err;
		case xltypeMulti:
			return rows(x) != rows(y) ? rows(x) <=> rows(y)
				: columns(x) != columns(y) ? columns(x) <=> columns(y)
				: std::lexicographical_compare_three_way(span(x).begin(), span(x).end(), span(y).begin(), span(y).end(), xll::compare);
		case xltypeInt:
			return x.val.w <=> y.val.w;
		case xltypeSRef:
			return SRef(x) <=> SRef(y);
		case xltypeRef:
			return x.val.mref.idSheet != y.val.mref.idSheet ? x.val.mref.idSheet <=> y.val.mref.idSheet
				: std::lexicographical_compare_three_way(ref(x).begin(), ref(x).end(), ref(y).begin(), ref(y).end());
		case xltypeBigData:
			return count(x) != count(y) ? count(x) <=> count(y)
				: std::lexicographical_compare_three_way(blob(x).begin(), blob(x).end(), blob(y).begin(), blob(y).end());
		}

		return std::partial_ordering::equivalent;
	}
} // namespace xll

constexpr auto operator<=>(const XLOPER12& x, const XLOPER12& y)
{
	return xll::compare(x, y);
}
constexpr bool operator==(const XLOPER12& x, const XLOPER12& y)
{
	return xll::compare(x, y) == 0;
}

#ifdef _DEBUG
static_assert((xll::Num(1.23) <=> xll::Num(1.23)) == 0);
static_assert((xll::Num(1.23) <=> xll::Num(1.24)) < 0);
static_assert((xll::Num(1.25) <=> xll::Num(1.24)) > 0);
#endif // _DEBUG

