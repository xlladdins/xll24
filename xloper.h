// xloper.h - XLOPER12 helpers
#pragma once
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
	constexpr int size(const XLOPER12& x) noexcept
	{
		return rows(x) * columns(x);
	}

	constexpr std::wstring_view view(const XLOPER12& x)
	{
		ensure(type(x) == xltypeStr);

		return std::wstring_view(x.val.str + 1, x.val.str[0]);
	}

	constexpr auto span(const XLOPER12& x)
	{
		ensure(type(x) == xltypeMulti);

		return std::span(x.val.array.lparray, size(x));
	}

	constexpr auto ref(const XLOPER12& x)
	{
		ensure(type(x) == xltypeRef);

		return std::span(x.val.mref.lpmref->reftbl, x.val.mref.lpmref->count);
	}

	constexpr auto blob(const XLOPER12& x)
	{
		ensure(type(x) == xltypeBigData);

		return std::span(x.val.bigdata.h.lpbData, x.val.bigdata.cbData);
	}

	constexpr bool equal(const XLOPER12& x, const XLOPER12& y)
	{
		int xtype = type(x);
		int ytype = type(y);

		if (xtype != ytype) {
			return false;
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
			return rows(x) == rows(y) && columns(x) == columns(y) 
				&& std::equal(span(x).begin(), span(x).end(), span(y).begin(), span(y).end(), xll::equal);
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

constexpr bool operator==(const XLOPER12& x, const XLOPER12& y)
{
	return xll::equal(x, y);
}

#ifdef _DEBUG
static_assert(xll::equal(xll::Num(1.23), xll::Num(1.23)));
static_assert(xll::Num(1.23) == xll::Num(1.23));
#endif // _DEBUG
