// xlref.h - XLREF type
#pragma once
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include "XLCALL.H"
#include "ensure.h"

namespace xll {

	struct REF : public XLREF12 {
		constexpr REF(const XLREF12& r) noexcept
			: XLREF12(r)
		{ }
		constexpr REF(int r, int c, int w = 1, int h = 1) noexcept
			: XLREF12{ .rwFirst = r, .rwLast = r + w - 1, .colFirst = c, .colLast = c + h - 1 }
		{ }
		constexpr REF(const XLOPER12& x) noexcept
			: REF(x.val.sref.ref)
		{ }
		
		constexpr bool operator==(const REF& r) const
		{
			return rwFirst == r.rwFirst
				&& rwLast == r.rwLast
				&& colFirst == r.colFirst
				&& colLast == r.colLast;
		}
		
	};
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

} // namespace xll

constexpr bool operator==(const XLREF12& r, const XLREF12& s)
{
	return r == s;
}
#ifdef _DEBUG
static_assert(xll::REF(1, 2).operator==(xll::REF(1, 2)));
static_assert(xll::REF(1, 2, 3, 4) == xll::REF(1, 2, 3, 4));
static_assert(xll::REF(1, 2, 3, 4) != xll::REF(1, 2, 3, 5));
#endif // _DEBUG
