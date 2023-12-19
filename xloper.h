#pragma once
#ifdef _DEBUG
#include <cassert>
#endif // _DEBUG
#include <algorithm>
#include <compare>
#include <span>
#include <string_view>
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include "XLCALL.H"
#include "ensure.h"

namespace xll {

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

	constexpr XLOPER12 Num(double num)
	{
		return XLOPER12{ .val = {.num = num}, .xltype = xltypeNum };
	}
	static_assert(type(Num(1.23)) == xltypeNum);

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
	/*
	constexpr XLOPER12 SRef(int rwFirst, int rwLast, int colFirst, int colLast)
	{
		return XLOPER12{ .val = {.sref = {.rwFirst = rwFirst, .rwLast = rwLast, .colFirst = colFirst, .colLast = colLast}}, .xltype = xltypeSRef };
	}
	*/
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

	constexpr std::wstring_view view(const XLOPER12& x)
	{
		return std::wstring_view(x.val.str + 1, x.val.str[0]);
	}
	static_assert(type(Str(L"\03" L"abc")) == xltypeStr);
	static_assert(view(Str(L"\03" L"abc")) == L"abc");

	template<size_t N = std::dynamic_extent>
	constexpr std::span<XLOPER12> span(const XLOPER12& x)
	{
		if (type(x) == xltypeMulti) {
			return std::span<XLOPER12>(x.val.array.lparray, x.val.array.rows * x.val.array.columns);
		}
		else {
			return std::span<XLOPER12>(&x, 1);
		}
	}

	// strong double comparison
	auto compare(double x, double y)
	{
		union u {
			double d;
			unsigned long long ull;
		} ux = { .d = x }, uy = { .d = y };

		return ux.ull <=> uy.ull;
	}
	//static_assert(compare(1.23, 1.23) == 0);

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
			const auto xspan = span(x);
			const auto yspan = span(y);
			return std::lexicographical_compare_three_way(
				xspan.begin(), xspan.end(),
				yspan.begin(), yspan.end());
				*/
		case xltypeMissing:
			return 0 <=> 0;
		case xltypeNil:
			return 0 <=> 0;
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


#ifdef _DEBUG
	static int test()
	{
		{

		}

		return 0;
	}
#endif // _DEBUG
	/*
	#ifdef _DEBUG
	static_assert(Str12().xltype == xltypeStr); // NULL string
	static_assert(Str12().val.str == nullptr);
	static_assert(Str12(L"").xltype == xltypeStr); // zero length string
	static_assert(Str12(L"").val.str[0] == 0);
	static_assert(Str12(L"\03" L"abc").xltype == xltypeStr);
	static_assert(Str12(L"\03" L"abc").val.str[0] == 3);
	static_assert(Str12(L"\03" L"abc").val.str[3] == L'c');

	static_assert(Str12() <=> Str12() == 0);
	static_assert(Str12(L"") <=> Str12(L"") == 0);
	static_assert((Str12(L"\001a") <=> Str12(L"\001b")) < 0);
	static_assert((Str12(L"\001a") <=> Str12(L"\001a")) == 0);
	static_assert((Str12(L"\001b") <=> Str12(L"\001a")) > 0);
	static_assert((Str12(L"\002aa") <=> Str12(L"\001a")) > 0);
	#endif // _DEBUG
	*/

	struct OPER12 : public XLOPER12 {
		// Nil
		constexpr OPER12() noexcept
			: XLOPER12{ xll::Nil() }
		{ }
		// Num
		constexpr explicit OPER12(double num) noexcept
			: XLOPER12{ .val = {.num = num}, .xltype = xltypeNum }
		{ }
		// Str
		constexpr OPER12(const XCHAR* str, XCHAR len)
			: XLOPER12{ .val = {.str = nullptr}, .xltype = xltypeStr }
		{
			if (str) {
				alloc(str, len);
			}
		}
		// NULL terminated string
		constexpr explicit OPER12(const XCHAR* str)
			: OPER12(str, static_cast<XCHAR>(wcslen(str)))
		{ }
		// Bool
		constexpr explicit OPER12(bool xbool)
			: XLOPER12{ .val = {.xbool = xbool}, .xltype = xltypeBool }
		{ }
		// Int
		constexpr explicit OPER12(int w) noexcept
			: XLOPER12{ .val = {.w = w}, .xltype = xltypeInt }
		{ }
		// Err
		constexpr explicit OPER12(xll::xlerr err) noexcept
			: XLOPER12{ .val = {.err = (int)err}, .xltype = xltypeErr }
		{ }
		// Multi
		constexpr OPER12(int r, int c)
		{
			alloc(r, c);
		}

		constexpr OPER12(const OPER12& o)
			: XLOPER12{ o }
		{
			if (xltype == xltypeStr) {
				new (this) OPER12(val.str + 1, val.str[0]);
			}
		}
		OPER12& operator=(const OPER12& o)
		{
			dealloc();
			if (type() == xltypeStr) {
				new (this) OPER12(o);
			}
			else {
				*this = o;
			}

			return *this;
		}
		OPER12(OPER12&& o) noexcept
			: XLOPER12{ .val = o.val , .xltype = std::exchange(o.xltype, xltypeNil) }
		{ }
		OPER12& operator=(OPER12&& o) noexcept
		{
			if (this != &o) {
				val = o.val;
				xltype = std::exchange(o.xltype, xltypeNil);
			}

			return *this;
		}
		~OPER12()
		{
			dealloc();
		}

		int type() const
		{
			return xll::type(*this);
		}

		int rows() const
		{
			return type() == xltypeMulti ? val.array.rows : 1;
		}
		int columns() const
		{
			return type() == xltypeMulti ? val.array.columns : 1;
		}
		int size() const
		{
			return rows() * columns();
		}

		const OPER12* begin() const
		{
			return type() == xltypeMulti ? (const OPER12*)val.array.lparray : this;
		}
		OPER12* begin()
		{
			return type() == xltypeMulti ? (OPER12*)val.array.lparray : this;
		}
		const OPER12* end() const
		{
			return type() == xltypeMulti ? (const OPER12*)val.array.lparray + size() : this + 1;
		}
		OPER12* end()
		{
			return type() == xltypeMulti ? (OPER12*)val.array.lparray + size() : this + 1;
		}

		const OPER12& operator[](int i) const
		{
			if (type() == xltypeMulti) {
				return (const OPER12&)val.array.lparray[i];
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
				for (auto& o : *this) {
					o.dealloc();
				}
				delete[] val.array.lparray;
			}
		}
	};
#ifdef _DEBUG
	//static_assert(OPER12().xltype == xltypeNil);
	//static_assert(OPER12(1.23).val.num == 1.23);
#endif // _DEBUG

} // namespace xll