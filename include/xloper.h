// xloper.h - XLOPER12 helpers
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <iostream>
#include <string_view>
#include "ref.h"

namespace xll {

	// return XLFree(x) in thread-safe functions
	// Freed by Excel when no longer needed.
	constexpr LPXLOPER12 XLFree(XLOPER12& x)
	{
		if (isAlloc(x)) {
			x.xltype |= xlbitXLFree;
		}

		return &x;
	}
	// return DLLFree(x) in thread-safe functions
	// Excel calls xlAutoFree12 when no longer needed.
	constexpr LPXLOPER12 DLLFree(XLOPER12& x)
	{
		if (isAlloc(x)) {
			x.xltype |= xlbitDLLFree;
		}

		return &x;
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
		}

		return 1;
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

	// STL friendly
	constexpr const XLOPER12* begin(const XLOPER12& x) noexcept
	{
		return isMulti(x) ? x.val.array.lparray : &x;
	}
	constexpr XLOPER12* begin(XLOPER12& x) noexcept
	{
		return isMulti(x) ? x.val.array.lparray : &x;
	}
	constexpr const XLOPER12* end(const XLOPER12& x) noexcept
	{
		return isMulti(x) ? x.val.array.lparray + size(x) : &x + 1;
	}
	constexpr XLOPER12* end(XLOPER12& x) noexcept
	{
		return isMulti(x) ? x.val.array.lparray + size(x) : &x + 1;
	}

	// Either one row or one column Multi
	constexpr bool isVector(const XLOPER12& x) noexcept
	{
		return isMulti(x) && (rows(x) == 1 || columns(x) == 1);
	}

	constexpr XLOPER12 index(const XLOPER12& x, int i) noexcept
	{
		return type(x) != xltypeMulti ? x : Multi(x)[i];
	}
	constexpr XLOPER12& index(XLOPER12& x, int i) noexcept
	{
		return type(x) != xltypeMulti ? x : Multi(x)[i];
	}
	constexpr XLOPER12 index(const XLOPER12& x, int i, int j) noexcept
	{
		return index(x, i * columns(x) + j);
	}
	constexpr XLOPER12& index(XLOPER12& x, int i, int j) noexcept
	{
		return index(x, i * columns(x) + j);
	}

	constexpr std::wstring_view view(const XLOPER12& x, size_t count) noexcept
	{
		return isStr(x) ? std::wstring_view(Str(x), count) : std::wstring_view();
	}
	constexpr std::wstring_view view(const XLOPER12& x) noexcept
	{
		return view(x, count(x));
	}

	constexpr auto span(const XLOPER12& x, size_t count)
	{
		return isMulti(x) ? std::span(Multi(x), count) : std::span<XLOPER12>();
	}
	constexpr auto span(const XLOPER12& x)
	{
		return span(x, count(x));
	}

	constexpr auto ref(const XLOPER12& x)
	{
		return isRef(x) ? std::span(Ref(x), count(x)) : std::span<XLREF12>();
	}

	constexpr auto blob(const XLOPER12& x)
	{
		ensure(isBigData(x));

		return std::span(BigData(x), count(x));
	}

	constexpr std::partial_ordering compare(const XLOPER12& x, const XLOPER12& y)
	{
		const int xtype = type(x);
		const int ytype = type(y);

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
			return x.val.err <=> y.val.err; // ??? always false like NaN
		case xltypeMulti:
			return rows(x) != rows(y) ? rows(x) <=> rows(y)
				: columns(x) != columns(y) ? columns(x) <=> columns(y)
				: std::lexicographical_compare_three_way(span(x).begin(), span(x).end(), span(y).begin(), span(y).end(), xll::compare);
		case xltypeInt:
			return x.val.w <=> y.val.w;
		case xltypeSRef:
			return x.val.sref.ref <=> y.val.sref.ref;
		case xltypeRef:
			return x.val.mref.idSheet != y.val.mref.idSheet ? x.val.mref.idSheet <=> y.val.mref.idSheet
				: std::lexicographical_compare_three_way(ref(x).begin(), ref(x).end(), ref(y).begin(), ref(y).end());
		case xltypeBigData:
			return count(x) != count(y) ? count(x) <=> count(y)
				: std::lexicographical_compare_three_way(blob(x).begin(), blob(x).end(), blob(y).begin(), blob(y).end());
		}

		return std::partial_ordering::equivalent;
	}

	// Lookup value corresponding to key in JSON like multi.
	// If s is 2x2 then it is row oriented.
	constexpr const XLOPER12 value(const XLOPER12& x, const XLOPER12& key)
	{
		if (rows(x) != 2 || columns(x) != 2) {
			return ErrNA;
		}

		if (rows(x) == 2) {
			for (int i = 0; i < columns(x); ++i) {
				if (xll::compare(index(x, 0, i), key) == std::partial_ordering::equivalent) {
					return index(x, 1, i);
				}
			}
		}
		else {
			for (int i = 0; i < rows(x); ++i) {
				if (xll::compare(index(x, i, 0), key) == std::partial_ordering::equivalent) {
					return index(x, i, 1);
				}
			}
		}

		return ErrNA;
	}
	
	// False-like value.
	constexpr bool isFalse(const XLOPER12& x)
	{
		return (isNum(x) && (Num(x) != Num(x) || Num(x) == 0))
			|| (isStr(x) && (count(x) == 0 || x.val.str[1] == 0))
			|| (isBool(x) && Bool(x) == false)
			|| (isRef(x) && count(x) == 0)
			|| isErr(x)
			|| (isMulti(x) && size(x) == 0)
			|| isMissing(x)
			|| isNil(x)
			|| (isSRef(x) && size(x) == 0)
			|| (isInt(x) && Int(x) == 0)
			|| (isBigData(x) && count(x) == 0)
			;
	}
	// True-like value
	constexpr bool isTrue(const XLOPER12& x)
	{
		return !isFalse(x);
	}
#ifdef _DEBUG
	static_assert(isFalse(Num(0)));
	static_assert(isTrue(Num(1)));
	static_assert(isFalse(Num(std::numeric_limits<double>::quiet_NaN())));
	static_assert(isFalse(PStr(L"")));
	static_assert(isFalse(PStr(L"\0x1")));
	static_assert(isFalse(Bool(false)));
	static_assert(isTrue(Bool(true)));
	static_assert(isFalse(ErrNA));
	static_assert(isFalse(XLOPER12{ .val = {.array = {.lparray = 0, .rows = 0, .columns = 0}}, .xltype = xltypeMulti }));
	static_assert(isFalse(XLOPER12{ .val = {.array = {.lparray = 0, .rows = 1, .columns = 0}}, .xltype = xltypeMulti }));
	static_assert(isFalse(XLOPER12{ .val = {.array = {.lparray = 0, .rows = 0, .columns = 1}}, .xltype = xltypeMulti }));
	static_assert(isTrue(XLOPER12{ .val = {.array = {.lparray = 0, .rows = 1, .columns = 1}}, .xltype = xltypeMulti }));
	static_assert(isFalse(Missing));
	static_assert(!isTrue(Nil));
	static_assert(isFalse(SRef(REF(0, 0, 0, 0))));
	static_assert(isFalse(SRef(REF(0, 0, 1, 0))));
	static_assert(isFalse(SRef(REF(0, 0, 0, 1))));
	static_assert(isTrue(SRef(REF(0, 0, 1, 1))));
	static_assert(isFalse(Int(0)));
	static_assert(isTrue(Int(1)));
	//static_assert(isFalse(BigData(nullptr, 0)));
#endif // _DEBUG


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

// Print XLOPER12 to stream.
// Use xlfEvaluate to convert a string to an OPER.
template<class T>
	requires xll::is_char<T>::value
inline std::basic_ostream<T>& operator<<(std::basic_ostream<T>& os, const XLOPER12& o)
{
	if (xll::isNum(o)) {
		os << o.val.num;
	}
	else if (xll::isStr(o)) {
		os << L"\"" << xll::view(o) << L"\"";
	}
	else if (xll::isBool(o)) {
		os << (o.val.xbool ? L"TRUE" : L"FALSE");
	}
	// isRef
	else if (xll::isErr(o)) {
		os << xll::xlerr_wstring(xll::xlerr{ o.val.err });
	}
	else if (xll::isMulti(o)) {
		os << L"{";
		for (int i = 0; i < xll::rows(o); ++i) {
			if (i > 0) {
				os << L";";
			}
			for (int j = 0; j < xll::columns(o); ++j) {
				if (j > 0) {
					os << L",";
				}
				os << xll::index(o, i, j);
			}
		}
		os << L"}";
	}
	// isSRef: use xlfReftext
	else if (xll::isMissing(o) || xll::isNil(o)) {
		os << xll::Empty;
	}
	else if (xll::isInt(o)) {
		os << o.val.w;
	}
	// isBigData
	else {
		os << L"<unknown>";
	}

	return os;
}

