// xloper.h - XLOPER12 helpers
#pragma once
#include "ref.h"

namespace xll {

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
			if (x.val.str && y.val.str) {
				return std::equal(
					view(x).begin(), view(x).end(),
					view(y).begin(), view(y).end());
			}
			else if (x.val.str) {
				return false;
			}
			else if (y.val.str) {
				return false;
			}
			else {
				return true;
			}
		case xltypeBool:
			return x.val.xbool == y.val.xbool;
		case xltypeErr:
			return x.val.err == y.val.err;
		case xltypeMulti:
			return std::equal(
				span(x).begin(), span(x).end(),
				span(y).begin(), span(y).end(), xll::equal);
		case xltypeInt:
			return x.val.w == y.val.w;
		case xltypeSRef:
			return x.val.sref.ref == y.val.sref.ref;
		case xltypeRef:
			return x.val.mref.idSheet == y.val.mref.idSheet
				&& x.val.mref.lpmref->count == y.val.mref.lpmref->count
				&& std::equal(
					x.val.mref.lpmref->reftbl, x.val.mref.lpmref->reftbl + x.val.mref.lpmref->count,
					y.val.mref.lpmref->reftbl, y.val.mref.lpmref->reftbl + y.val.mref.lpmref->count);
		default:
			return true;
		}
	}
	static_assert(equal(Num(1.23), Num(1.23)));
	constexpr bool operator==(const XLOPER12& x, const XLOPER12& y)
	{
		return equal(x, y);
	}
	static_assert(Num(1.23) == Num(1.23));

} // namespace xll
