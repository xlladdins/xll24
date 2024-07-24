// ref.h - REF class to construct XLREF12 type
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <compare>
#include "defines.h"

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

	// Upper left and lower right corners of single reference. 
	struct REF : public XLREF12 {
		constexpr REF(const XLREF12& r) noexcept
			: XLREF12(r)
		{ }
		// Upper left corner and optional height and width.
		constexpr REF(int r, int c, int h = 1, int w = 1) noexcept
			: XLREF12{ .rwFirst = r, .rwLast = r + h - 1, .colFirst = c, .colLast = c + w - 1 }
		{ }
		REF(const REF&) noexcept = default;
		REF& operator=(const REF&) noexcept = default;
		~REF() noexcept = default;
	};

	constexpr REF reshape(const XLREF12& r, int h, int w)
	{
		return REF(r.rwFirst, r.colFirst, h, w);
	}
	constexpr REF translate(const XLREF12& r, int dr, int dc)
	{
		return REF(r.rwFirst + dr, r.colFirst + dc, rows(r), columns(r));
	}

} // namespace xll

constexpr bool operator==(const XLREF12& r, const XLREF12& s)
{
	return r.rwFirst == s.rwFirst
		&& r.rwLast == s.rwLast
		&& r.colFirst == s.colFirst
		&& r.colLast == s.colLast;
}
// Upper left corner lexicographically less.
constexpr bool operator<(const XLREF12 & r, const XLREF12 & s)
{
	return r.rwFirst < s.rwFirst
		|| (r.rwFirst == s.rwFirst && r.colFirst < s.colFirst);
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
static_assert(reshape(xll::REF(1, 2, 3, 4), 2, 3) == xll::REF(1, 2, 2, 3));
static_assert(translate(xll::REF(1, 2, 3, 4), 1, 1) == xll::REF(2, 3, 3, 4));
#endif // _DEBUG
