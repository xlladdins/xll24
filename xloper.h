#pragma once
#ifdef _DEBUG
#include <cassert>
#endif // _DEBUG
#include <algorithm>
#include <bit>
#include <compare>
#include <span>
#include <string_view>
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include "XLCALL.H"
#include "ensure.h"

namespace xll {

	// strong double comparison
	constexpr auto compare(double x, double y)
	{
		int64_t lx = std::bit_cast<int64_t>(x);
		int64_t ly = std::bit_cast<int64_t>(y);

		return lx <=> ly;
	}
#ifdef _DEBUG
	static_assert(compare(1.23, 1.23) == 0);
	static_assert(compare(-1.23, 1.23) < 0);
	static_assert(compare(1.23, -1.23) > 0);
#endif // _DEBUG
	constexpr int xltypeNull = xltypeNil | xltypeMissing;
	constexpr int xltypeScalar = xltypeNum | xltypeBool | xltypeErr | xltypeSRef | xltypeInt | xltypeNull;
	constexpr int xlbitFree = xlbitDLLFree | xlbitXLFree;

	//!!! X-macro
	enum class xlerr {
		Null = xlerrNull,
		Div0 = xlerrDiv0,
		Value = xlerrValue,
		Ref = xlerrRef,
		Name = xlerrName,
		Num = xlerrNum,
		NA = xlerrNA,
		GettingData = xlerrGettingData,
	};

	constexpr int type(const XLOPER12& x)
	{
		return x.xltype & ~(xlbitFree);
	}

	constexpr int rows(const XLOPER12& x)
	{
		switch (type(x)) {
		case xltypeMulti:
			return x.val.array.rows;
		case xltypeSRef:
			return x.val.sref.ref.rwLast - x.val.sref.ref.rwFirst + 1;
		case xltypeMissing:
		case xltypeNil:
			return 0;
		default:
			return 1;
		}
	}
	constexpr int columns(const XLOPER12& x)
	{
		switch (type(x)) {
		case xltypeMulti:
			return x.val.array.columns;
		case xltypeSRef:
			return x.val.sref.ref.colLast - x.val.sref.ref.colFirst + 1;
		case xltypeMissing:
		case xltypeNil:
			return 0;
		default:
			return 1;
		}
	}
	constexpr auto size(const XLOPER12& x)
	{
		return rows(x) * columns(x);
	}

	constexpr std::wstring_view view(const XLOPER12& x)
	{
		return std::wstring_view(x.val.str + 1, x.val.str[0]);
	}

	constexpr std::span<XLOPER12> span(const XLOPER12& x)
	{
		return std::span<XLOPER12>(x.val.array.lparray, size(x));
	}

	constexpr XLOPER12 Num(double num)
	{
		return XLOPER12{ .val = {.num = num}, .xltype = xltypeNum };
	}
	static_assert(type(Num(1.23)) == xltypeNum);
	static_assert(Num(1.23).val.num == 1.23);

	constexpr XLOPER12 Bool(bool xbool)
	{
		return XLOPER12{ .val = {.xbool = xbool}, .xltype = xltypeBool };
	}
	static_assert(type(Bool(true)) == xltypeBool);

	constexpr XLOPER12 Err(xlerr err)
	{
		return XLOPER12{ .val = {.err = (int)err}, .xltype = xltypeErr };
	}
	static_assert(type(Err(xlerr::Null)) == xltypeErr);

	constexpr XLOPER12 Missing()
	{
		return XLOPER12{ .xltype = xltypeMissing };
	}
	constexpr XLOPER12 Nil()
	{
		return XLOPER12{ .xltype = xltypeNil };
	}

	constexpr XLREF12 XLRef(int r, int c, int w = 1, int h = 1)
	{
		return XLREF12{ .rwFirst = r, .rwLast = r + w - 1, .colFirst = c, .colLast = c + h - 1};
	}
	constexpr XLOPER12 SRef(int r, int c, int w = 1, int h = 1)
	{
		return XLOPER12{ .val = {.sref = {.count = 1, .ref = XLRef(r, c, w, h)}}, .xltype = xltypeSRef};
	}
	static_assert(type(SRef(2, 3)) == xltypeSRef); // B3
	static_assert(rows(SRef(2, 3)) == 1);
	static_assert(columns(SRef(2, 3)) == 1);
	static_assert(size(SRef(2, 3)) == 1);
	static_assert(type(SRef(2, 3, 4, 5)) == xltypeSRef); // B3:E8
	static_assert(rows(SRef(2, 3, 4, 5)) == 4);
	static_assert(columns(SRef(2, 3, 4, 5)) == 5);
	static_assert(size(SRef(2, 3, 4, 5)) == 20);

	constexpr XLOPER12 Int(int w)
	{
		return XLOPER12{ .val = {.w = w}, .xltype = xltypeInt };
	}
	static_assert(type(Int(123)) == xltypeInt);

	// counted string
	constexpr XLOPER12 Str(const XCHAR* str = L"\000")
	{
		return XLOPER12{ .val = {.str = const_cast<wchar_t*>(str)}, .xltype = xltypeStr };
	}
	static_assert(type(Str()) == xltypeStr);
	static_assert(Str().val.str[0] == 0);

	constexpr XLOPER12 Multi(int r, int c)
	{
		return XLOPER12{ .val = {.array = {.lparray = nullptr, .rows = r, .columns = c}}, .xltype = xltypeMulti };
	}
	static_assert(type(Multi(2, 3)) == xltypeMulti);
	static_assert(rows(Multi(2, 3)) == 2);
	static_assert(columns(Multi(2, 3)) == 3);
	static_assert(size(Multi(2, 3)) == 6);

	static_assert(type(Str(L"\03" L"abc")) == xltypeStr);
	static_assert(view(Str(L"\03" L"abc")) == L"abc");


	constexpr auto operator<=>(const XLOPER12& x, const XLOPER12& y)
	{
		int xtype = type(x);
		int ytype = type(y);

		if (xtype != ytype) {
			return xtype <=> ytype;
		}

		switch (xtype) {
		case xltypeNum:
			return compare(x.val.num, y.val.num);
		case xltypeStr:
			if (x.val.str && y.val.str) {
				return std::lexicographical_compare_three_way(
					view(x).begin(), view(x).end(),
					view(y).begin(), view(y).end());
			}
			else if (x.val.str) {
				return 1 <=> 0;
			}
			else if (y.val.str) {
				return -1 <=> 0;
			}
			else {
				return 0 <=> 0;
			}
		case xltypeBool:
			return x.val.xbool <=> y.val.xbool;
		case xltypeErr:
			return x.val.err <=> y.val.err;
		case xltypeMulti:
			return 0 <=> 0;
			/*
			return std::lexicographical_compare_three_way(
				span(x).begin(), span(x).end(),
				span(y).begin(), span(y).end());
				*/
		case xltypeInt:
			return x.val.w <=> y.val.w;
		case xltypeSRef:
			return 0 <=> 0;
		case xltypeRef:
			return 0 <=> 0;
		default:
			return 0 <=> 0;
		}
	}
	static_assert(Num(1.23) <=> Num(1.23) == 0);

	struct OPER : public XLOPER12 {
		// Nil
		constexpr OPER() noexcept
			: XLOPER12{ Nil() }
		{ }
		// Num
		constexpr explicit OPER(double num) noexcept
			: XLOPER12{ Num(num) }
		{ }
		// Str
		constexpr OPER(const XCHAR* str, XCHAR len)
			: XLOPER12{ Str() }
		{
			if (str) {
				alloc(str, len);
			}
		}
		// NULL terminated string
		constexpr explicit OPER(const XCHAR* str)
			: OPER(str, static_cast<XCHAR>(wcslen(str)))
		{ }
		// Bool
		constexpr explicit OPER(bool xbool)
			: XLOPER12{ Bool(xbool)}
		{ }
		// Int
		constexpr explicit OPER(int w) noexcept
			: XLOPER12{ Int(w)}
		{ }
		// Err
		constexpr explicit OPER(xll::xlerr err) noexcept
			: XLOPER12{ Err(err) }
		{ }
		// Multi
		constexpr OPER(int r, int c, const XLOPER12* a)
			: XLOPER12{ Multi(r, c) }
		{
			alloc(r, c);
			if (a) {
				ensure(r * c == xll::size(*a))
				std::copy_n(a, r * c, val.array.lparray);
			}
		}
		constexpr OPER(std::initializer_list<XLOPER12> a)
			: OPER(static_cast<int>(a.size()), 1, a.begin())
		{ }
		OPER& reshape(int r, int c)
		{
			ensure(r * c == size());
			val.array.rows = r;
			val.array.columns = c;

			return *this;
		}

		constexpr OPER(const OPER& o)
			: XLOPER12{ o }
		{
			if (xltype == xltypeStr) {
				new (this) OPER(val.str + 1, val.str[0]);
			}
		}
		OPER& operator=(const OPER& o)
		{
			dealloc();
			if (type() == xltypeStr) {
				new (this) OPER(o);
			}
			else {
				*this = o;
			}

			return *this;
		}
		OPER(OPER&& o) noexcept
			: XLOPER12{ .val = o.val , .xltype = std::exchange(o.xltype, xltypeNil) }
		{ }
		OPER& operator=(OPER&& o) noexcept
		{
			if (this != &o) {
				val = o.val;
				xltype = std::exchange(o.xltype, xltypeNil);
			}

			return *this;
		}
		~OPER()
		{
			dealloc();
		}

		constexpr int type() const
		{
			return xll::type(*this);
		}
		constexpr int size() const
		{
			return xll::size(*this);
		}
		constexpr std::span<OPER> span()
		{
			return std::span((OPER*)val.array.lparray, size());
		}

		OPER& operator[](int i)
		{
			if (type() == xltypeMulti) {
				return (OPER&)val.array.lparray[i];
			}
			else {
				ensure(i == 0);
				return *this;
			}
		}
		const OPER& operator[](int i) const
		{
			if (type() == xltypeMulti) {
				return (const OPER&)val.array.lparray[i];
			}
			else {
				ensure(i == 0);
				return *this;
			}
		}
	private:
		constexpr void alloc(const XCHAR* str, XCHAR len)
		{
			val.str = const_cast<wchar_t*>(new wchar_t[1 + len]);
			val.str[0] = len;
			std::copy_n(str, len, val.str + 1);

		}
		constexpr void alloc(int r, int c)
		{
			xltype = xltypeMulti;
			val.array.rows = r;
			val.array.columns = c;
			val.array.lparray = nullptr;
			if (r && c) {
				val.array.lparray = new XLOPER12[r * c];
			}
		}
		void dealloc()
		{
			if (xltype & xlbitXLFree) {
				::Excel12v(xlFree, 0, 1, (LPXLOPER12*)this);
			}
			else if (xltype == xltypeStr) {
				delete[] val.str;
			}
			else if (xltype == xltypeMulti) {
				for (auto& o : span()) {
					o.dealloc();
				}
				delete[] val.array.lparray;
			}
		}
	
	};
#ifdef _DEBUG
	//static_assert(Nil().xltype == xltypeNil);
	//static_assert(Num(1.23).val.num == 1.23);
#endif // _DEBUG
#ifdef _DEBUG
	static int test()
	{
		{
			OPER o;
			assert(o.xltype == xltypeNil);
			OPER o2{ o };
			//assert(o == o2);

		}
		{
			static_assert(Str().xltype == xltypeStr); // zero length string
			static_assert(Str().val.str[0] == 0);
			static_assert(Str(L"").xltype == xltypeStr);
			static_assert(Str(L"").val.str[0] == 0);
			static_assert(Str(L"\03" L"abc").xltype == xltypeStr);
			static_assert(Str(L"\03" L"abc").val.str[0] == 3);
			static_assert(Str(L"\03" L"abc").val.str[3] == L'c');

			static_assert(Str() <=> Str() == 0);
			static_assert(Str(L"") <=> Str(L"") == 0);
			static_assert((Str(L"\001a") <=> Str(L"\001b")) < 0);
			static_assert((Str(L"\001a") <=> Str(L"\001a")) == 0);
			static_assert((Str(L"\001b") <=> Str(L"\001a")) > 0);
			static_assert((Str(L"\002aa") <=> Str(L"\001a")) > 0);
		}
		{

		}

		return 0;
	}
#endif // _DEBUG

} // namespace xll