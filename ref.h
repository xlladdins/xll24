// ref.h - REF class to construct XLREF12 type
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <compare>
#include "defines.h"
#include "ensure.h"

namespace xll {

	constexpr int rows(const XLREF12& r)
	{
		return r.rwLast - r.rwFirst + 1;
	}
	constexpr int columns(const XLREF12& r)
	{
		return r.colLast - r.colFirst + 1;
	}
	constexpr int size(const XLREF12& r)
	{
		return rows(r) * columns(r);
	}

	// Upper left and lower right corners of reference. 
	struct REF : public XLREF12 {
		constexpr REF(const XLREF12& r) noexcept
			: XLREF12(r)
		{ }
		// Upper left corner and optional width and height.
		constexpr REF(int r, int c, int w = 1, int h = 1) noexcept
			: XLREF12{ .rwFirst = r, .rwLast = r + w - 1, .colFirst = c, .colLast = c + h - 1 }
		{ }
		// Construct from XLPOPER12.
		constexpr REF(const XLOPER12& x)
			: REF(x.val.sref.ref)
		{
			ensure(x.xltype == xltypeSRef);
		}
#ifdef _DEBUG
#define XLL_SREF1234 {.rwFirst = 1, .rwLast = 2, .colFirst = 3, .colLast = 4 }
#define XLL_XLOPER1234 XLOPER12{ .val = {.sref = {.ref = XLL_SREF1234 }} , .xltype = xltypeSRef }
		static_assert(XLL_XLOPER1234.xltype == xltypeSRef);
		static_assert(SRef(XLL_XLOPER1234).rwFirst == 1);
		static_assert(SRef(XLL_XLOPER1234).rwLast == 2);
		static_assert(SRef(XLL_XLOPER1234).colFirst == 3);
		static_assert(SRef(XLL_XLOPER1234).colLast == 4);
#endif // _DEBUG

	};

} // namespace xll

constexpr bool operator==(const XLREF12& r, const XLREF12& s)
{
	return r.rwFirst == s.rwFirst
		&& r.rwLast == s.rwLast
		&& r.colFirst == s.colFirst
		&& r.colLast == s.colLast;
}
// r is strictly contained in s
constexpr bool operator<(const XLREF12 & r, const XLREF12 & s)
{
	return r.rwFirst > s.rwFirst
		&& r.colFirst > s.colFirst
		&& r.rwLast < s.rwLast
		&& r.colLast < s.colLast;
}
constexpr auto operator<=>(const XLREF12& r, const XLREF12& s)
{
	return r < s ? std::strong_ordering::less
		 : s < r ? std::strong_ordering::greater
		 : std::strong_ordering::equivalent;
}

#ifdef _DEBUG
static_assert(xll::REF(1, 2) == xll::REF(1, 2));
static_assert(xll::REF(1, 2, 3, 4) == xll::REF(1, 2, 3, 4));
static_assert(xll::REF(1, 2, 3, 4) != xll::REF(1, 2, 3, 5));
static_assert(rows(xll::REF(1, 2, 2, 3)) == 2);
static_assert(columns(xll::REF(1, 2, 2, 3)) == 3);
static_assert(size(xll::REF(1, 2, 2, 3)) == 6);
static_assert(xll::REF(XLL_XLOPER1234).colFirst == 3);
static_assert(xll::REF(XLL_XLOPER1234).colLast == 4);
//static_assert(xll::REF(XLL_XLOPER1234) == xll::REF(1, 2, 2, 3));
#undef XLL_SREF1234
#undef XLL_XLOPER1234
#endif // _DEBUG
