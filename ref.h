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
		REF(const REF& o) noexcept = default;
		REF& operator=(const REF& o) noexcept = default;
		~REF() noexcept = default;
	};

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
#endif // _DEBUG
