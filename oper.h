#pragma once
#ifdef _DEBUG
#include <cassert>
#endif // _DEBUG
#include <initializer_list>
#include "xloper.h"
#include "utf8.h"

namespace xll {

	struct OPER : public XLOPER12 {
		// Nil
		constexpr OPER() noexcept
			: XLOPER12{ Nil() }
		{ }
		// Num
		constexpr explicit OPER(double num) noexcept
			: XLOPER12{ Num(num) }
		{ }
		OPER& operator=(double num) noexcept
		{
			dealloc();
			*this = Num(num);

			return *this;
		}
		constexpr bool operator==(double num) const
		{
			return type(*this) == xltypeNum && val.num == num;
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
			: OPER(str, str ? static_cast<XCHAR>(wcslen(str)) : 0)
		{ }
		OPER& operator=(const XCHAR* str)
		{
			dealloc();
			*this = OPER(str);

			return *this;
		}
		constexpr bool operator==(const XCHAR* str) const
		{
			return type(*this) == xltypeStr && view(*this) == str;
		}
		explicit OPER(const char* str)
			: XLOPER12{ .val = {.str = utf8::mbstowcs(str)}, .xltype = xltypeStr }
		{ }

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
			: XLOPER12{ Err(static_cast<int>(err)) }
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
			ensure(r * c == size(*this));
			val.array.rows = r;
			val.array.columns = c;

			return *this;
		}
		// Convert to 1 x 1 multi.
		OPER& enlist()
		{
			if (type(*this) != xltypeMulti) {
				OPER o(*this);
				dealloc();
				alloc(1, 1, &o);
			}

			return *this;
		}
		// key-value pairs
		OPER(std::initializer_list<std::pair<const XCHAR*, OPER>> o)
		{
			alloc(2, (int)o.size(), nullptr);
			int i = 0;
			for (auto [k, v] : o) {
				operator()(0, i) = k;
				operator()(1, i) = v;
				++i;
			}
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

		constexpr std::span<OPER> span() noexcept
		{
			return std::span((OPER*)val.array.lparray, size(*this));
		}
		constexpr std::span<OPER> row(int r) noexcept
		{
			return std::span((OPER*)val.array.lparray + r * columns(*this), columns(*this));
		}

		OPER& operator[](int i)
		{
			if (type(*this) == xltypeMulti) {
				return (OPER&)val.array.lparray[i];
			}
			else {
				ensure(i == 0);
				return *this;
			}
		}
		const OPER& operator[](int i) const
		{
			if (type(*this) == xltypeMulti) {
				return (const OPER&)val.array.lparray[i];
			}
			else {
				ensure(i == 0);
				return *this;
			}
		}

		OPER& operator()(int i, int j)
		{
			if (type(*this) == xltypeMulti) {
				return (OPER&)val.array.lparray[i * columns(*this) + j];
			}
			else {
				ensure(i == 0 && j == 0);
				return *this;
			}
		}
		const OPER& operator()(int i, int j) const
		{
			if (type(*this) == xltypeMulti) {
				return (const OPER&)val.array.lparray[i * columns(*this) + j];
			}
			else {
				ensure(i == 0 && j == 0);
				return *this;
			}
		}

		// Find index of string in first row
		constexpr int index(const XCHAR* str) const
		{
			int i = 0;

			while (i < columns(*this) && view(operator()(0, i)) != str) {
				++i;
			}

			return i;
		}
		// JSON like object with keys in first row.
		constexpr bool is_object() const
		{
			if (type(*this) != xltypeMulti || rows(*this) != 2) {
				return false;
			}

			for (int i = 0; i < columns(*this); ++i) {
				if (type(operator()(0, i)) != xltypeStr) {
					return false;
				}
			}

			return true;
		}
		OPER& operator[](const XCHAR* str)
		{
			ensure(is_object());
			int i = index(str);
			ensure(i < columns(*this));

			return operator()(1, i);
		}
		const OPER& operator[](const XCHAR* str) const
		{
			ensure(is_object());
			int i = index(str);
			ensure(i < columns(*this));

			return operator()(1, i);
		}

	private:
		// Str
		constexpr void alloc(const XCHAR* str, XCHAR len)
		{
			xltype = xltypeStr;
			val.str = new wchar_t[1 + static_cast<size_t>(len)];
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
				val.array.lparray = new XLOPER12[(size_t)r * c];
				for (int i = 0; i < r * c; ++i) {
					new (val.array.lparray + i) OPER(a ? a[i] : OPER{});
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
#endif // _DEBUG

} // namespace xll

// Concatenate strings
inline xll::OPER operator&(const XLOPER12& x, const XLOPER12& y)
{
	ensure(xll::type(x) == xltypeStr && xll::type(y) == xltypeStr);
	
	std::wstring xy(xll::view(x));
	xy.append(xll::view(y));

	return xll::OPER(xy.c_str(), (XCHAR)xy.size());
}
