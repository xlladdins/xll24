// ref.h - REF class to construct XLREF12 type
#pragma once
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

	struct REF : public XLREF12 {
		constexpr REF(const XLREF12& r) noexcept
			: XLREF12(r)
		{ }
		constexpr REF(int r, int c, int w = 1, int h = 1) noexcept
			: XLREF12{ .rwFirst = r, .rwLast = r + w - 1, .colFirst = c, .colLast = c + h - 1 }
		{ }
		constexpr REF(const XLOPER12& x)
			: REF(x.val.sref.ref)
		{
			ensure(x.xltype == xltypeSRef);
		}
		
		constexpr bool operator==(const REF& r) const
		{
			return rwFirst == r.rwFirst
				&& rwLast == r.rwLast
				&& colFirst == r.colFirst
				&& colLast == r.colLast;
		}
		
	};

#ifdef _DEBUG
	static_assert(rows(REF(1,2,3,4)) == 3);
	static_assert(columns(REF(1, 2, 3, 4)) == 4);
	static_assert(size(REF(1, 2, 3, 4)) == 12);
#endif // _DEBUG

} // namespace xll

constexpr bool operator==(const XLREF12& r, const XLREF12& s)
{
	return xll::REF(r) == xll::REF(s);
}

#ifdef _DEBUG
static_assert(xll::REF(1, 2).operator==(xll::REF(1, 2)));
static_assert(xll::REF(1, 2, 3, 4) == xll::REF(1, 2, 3, 4));
static_assert(xll::REF(1, 2, 3, 4) != xll::REF(1, 2, 3, 5));
#endif // _DEBUG
