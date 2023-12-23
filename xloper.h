// xloper.h - XLOPER12 helpers
#pragma once
#include <span>
#include "ref.h"

namespace xll {

	constexpr int type(const XLOPER& x) noexcept
	{
		return x.xltype & ~(xlbitFree);
	}
	constexpr int type(const XLOPER12& x) noexcept
	{
		return x.xltype & ~(xlbitFree);
	}

	constexpr int rows(const XLOPER12& x) noexcept
	{
		switch (type(x)) {
		case xltypeMulti:
			return x.val.array.rows;
		case xltypeSRef:
			return rows(x.val.sref.ref);
		case xltypeMissing:
		case xltypeNil:
			return 0;
		default:
			return 1;
		}
	}
	constexpr int columns(const XLOPER12& x) noexcept
	{
		switch (type(x)) {
		case xltypeMulti:
			return x.val.array.columns;
		case xltypeSRef:
			return columns(x.val.sref.ref);
		case xltypeMissing:
		case xltypeNil:
			return 0;
		default:
			return 1;
		}
	}
	constexpr auto size(const XLOPER12& x) noexcept
	{
		return rows(x) * columns(x);
	}

	constexpr bool equal(const XLOPER12& x, const XLOPER12& y)
	{
		int xtype = type(x);
		int ytype = type(y);

		if (xtype != ytype) {
			return xtype == ytype;
		}

		switch (xtype) {
		case xltypeNum:
			return x.val.num == y.val.num;
		case xltypeStr:
			return std::equal(view(x).begin(), view(x).end(), view(y).begin(), view(y).end());
		case xltypeBool:
			return x.val.xbool == y.val.xbool;
		case xltypeErr:
			return x.val.err == y.val.err;
		case xltypeMulti:
			return std::equal(span(x).begin(), span(x).end(), span(y).begin(), span(y).end(), xll::equal);
		case xltypeInt:
			return x.val.w == y.val.w;
		case xltypeSRef:
			return SRef(x) == SRef(y);
		case xltypeRef:
			return x.val.mref.idSheet == y.val.mref.idSheet
				&& std::equal(ref(x).begin(), ref(x).end(), ref(y).begin(), ref(y).end());
		case xltypeBigData:
			return blob(x).size() == blob(y).size()
				&& std::equal(blob(x).begin(), blob(x).end(), blob(y).begin(), blob(y).end());
		default:
			return true;
		}
	}
} // namespace xll

static_assert(xll::equal(xll::Num(1.23), xll::Num(1.23)));
constexpr bool operator==(const XLOPER12& x, const XLOPER12& y)
{
	return xll::equal(x, y);
}
static_assert(xll::Num(1.23) == xll::Num(1.23));
