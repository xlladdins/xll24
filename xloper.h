#pragma once
#ifdef _DEBUG
#include <cassert>
#endif // _DEBUG
#include <algorithm>
#include <bit>
#include <compare>
#include <span>
#include <string_view>
#include "xlref.h"

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
			return rows(x.val.sref.ref);
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
			return columns(x.val.sref.ref);
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

	constexpr XLOPER12 Int(int w)
	{
		return XLOPER12{ .val = {.w = w}, .xltype = xltypeInt };
	}
	static_assert(type(Int(123)) == xltypeInt);

	// counted static string
	constexpr XLOPER12 Str(const XCHAR* str = L"\000")
	{
		return XLOPER12{ .val = {.str = const_cast<wchar_t*>(str)}, .xltype = xltypeStr };
	}
	static_assert(type(Str()) == xltypeStr);
	static_assert(Str().val.str[0] == 0);
	static_assert(type(Str(L"\003abc")) == xltypeStr);
	static_assert(Str(L"\003abc").val.str[0] == 3);
	static_assert(Str(L"\003abc").val.str[3] == L'c');

	constexpr XLOPER12 Multi(int r, int c)
	{
		return XLOPER12{ .val = {.array = {.lparray = nullptr, .rows = r, .columns = c}}, .xltype = xltypeMulti };
	}
	static_assert(type(Multi(2, 3)) == xltypeMulti);
	static_assert(rows(Multi(2, 3)) == 2);
	static_assert(columns(Multi(2, 3)) == 3);
	static_assert(size(Multi(2, 3)) == 6);

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
			return 0 == 0;
		default:
			return false;
		}
	}
	static_assert(equal(Num(1.23), Num(1.23)));
	constexpr bool operator==(const XLOPER12& x, const XLOPER12& y)
	{
		return equal(x, y);
	}
	static_assert(Num(1.23) == Num(1.23));

	struct OPER : public XLOPER12 {
		// Nil
		constexpr OPER() noexcept
			: XLOPER12{ Nil() }
		{ }
		// Num
		constexpr explicit OPER(double num) noexcept
			: XLOPER12{ Num(num) }
		{ }
		constexpr bool operator==(double num) const
		{
			return type() == xltypeNum && val.num == num;
		}
		// string and length
		constexpr OPER(const XCHAR* str, XCHAR len)
		{
			if (str) {
				alloc(str, len);
			}
		}
		// NULL terminated string
		constexpr explicit OPER(const XCHAR* str)
			: OPER(str, static_cast<XCHAR>(wcslen(str)))
		{ }
		constexpr bool operator==(const XCHAR* str) const
		{
			return type() == xltypeStr && view(*this) == str;
		}
		// Bool
		constexpr explicit OPER(bool xbool)
			: XLOPER12{ Bool(xbool) }
		{ }
		// Int
		constexpr explicit OPER(int w) noexcept
			: XLOPER12{ Int(w) }
		{ }
		// Err
		constexpr explicit OPER(xll::xlerr err) noexcept
			: XLOPER12{ Err(err) }
		{ }
		// Multi
		constexpr OPER(int r, int c, const XLOPER12* a)
		{
			alloc(r, c, a);
		}
		constexpr OPER(std::initializer_list<XLOPER12> a)
			: OPER(1, static_cast<int>(a.size()), a.begin())
		{ }
		OPER& reshape(int r, int c)
		{
			ensure(r * c == size());
			val.array.rows = r;
			val.array.columns = c;

			return *this;
		}

		constexpr OPER(const XLOPER12& o)
			: XLOPER12{ o }
		{
			if (xll::type(o) == xltypeStr) {
				alloc(val.str + 1, val.str[0]);
			}
			else if (xll::type(o) == xltypeMulti) {
				alloc(rows(o), columns(o), o.val.array.lparray);
			}
		}
		constexpr OPER(const OPER& o)
			: OPER(static_cast<const XLOPER12&>(o))
		{ }

		OPER& operator=(const XLOPER12& x)
		{
			if (this != &x) {
				OPER o(x);
				dealloc();
				*this = std::move(o);
			}

			return *this;
		}
		OPER& operator=(const OPER& o)
		{
			return operator=(static_cast<const XLOPER12&>(o));
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

		OPER& operator()(int i, int j)
		{
			if (type() == xltypeMulti) {
				return (OPER&)val.array.lparray[i * columns(*this) + j];
			}
			else {
				ensure(i == 0 && j == 0);
				return *this;
			}
		}
		const OPER& operator()(int i, int j) const
		{
			if (type() == xltypeMulti) {
				return (const OPER&)val.array.lparray[i * columns(*this) + j];
			}
			else {
				ensure(i == 0 && j == 0);
				return *this;
			}
		}
	private:
		// Str
		constexpr void alloc(const XCHAR* str, XCHAR len)
		{
			xltype = xltypeStr;
			val.str = new wchar_t[1 + len];
			val.str[0] = len;
			std::copy_n(str, len, val.str + 1);
		}
		// Multi
		constexpr void alloc(int r, int c, const XLOPER12* a)
		{
			xltype = xltypeMulti;
			val.array.rows = r;
			val.array.columns = c;
			if (r && c) {
				val.array.lparray = new XLOPER12[r * c];
				for (int i = 0; i < r * c; ++i) {
					new (val.array.lparray + i) OPER(a[i]);
				}
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

			xltype = xltypeNil;
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

			static_assert(Str() == Str());
			static_assert(Str(L"") == Str(L""));
			//static_assert((Str(L"\001a") == Str(L"\001b")) < 0);
			static_assert((Str(L"\001a") == Str(L"\001a")));
			//static_assert((Str(L"\001b") == Str(L"\001a")) > 0);
			//static_assert((Str(L"\002aa") == Str(L"\001a")) > 0);
		}
		{
			OPER m({ OPER(1.23), OPER(L"abc") });
			assert(type(m) == xltypeMulti);
			assert(rows(m) == 1);
			assert(columns(m) == 2);
			assert(size(m) == 2);
			assert(m[0] == 1.23);
			assert(m[1] == L"abc");
			m[0] = m;
			m[1] = m;
			m[1][1] = m;
			OPER m2{ m };
			assert(m == m2);
			m = m2;
			assert(!(m != m2));
		}

		return 0;
	}
#endif // _DEBUG

} // namespace xll