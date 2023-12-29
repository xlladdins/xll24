#pragma once
#ifdef _DEBUG
#include <cassert>
#endif // _DEBUG
#include <initializer_list>
#include "xloper.h"
#include "utf8.h"

namespace xll {

	// Length of null teriminated string.
	constexpr XCHAR len(const XCHAR* s)
	{
		return *s ? 1 + len(s + 1) : 0;
	}
	static_assert(len(L"abc") == 3);

	struct OPER : public XLOPER12 {
		// Nil is the default.
		constexpr OPER() noexcept
			: XLOPER12{ Nil }
		{ }

		// Num - 64-bit IEEE floating point
		constexpr explicit OPER(double num) noexcept
			: XLOPER12{ Num(num) }
		{ }
		OPER& operator=(double num) noexcept
		{
			dealloc();
			*this = Num(num);

			return *this;
		}
		constexpr bool operator==(double num) const noexcept
		{
			return type(*this) == xltypeNum && val.num == num;
		}
		double as_num() const noexcept
		{
			return xll::as_num(*this);
		}

		// Str - Counted wide character string.
		constexpr OPER(const XCHAR* str, XCHAR len)
		{
			if (str) {
				alloc(str, len);
			}
		}
		/*
		template<size_t len>
		constexpr OPER(const XCHAR (&str)[len])
			: OPER(str, static_cast<XCHAR>(len - 1))
		{ }
		*/
		// NULL terminated string
		constexpr explicit OPER(const XCHAR* str)
			: OPER(str, str ? len(str) : 0)
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
		// Distinguish from OPER(bool) constructor
		template<typename T>
		explicit OPER(T str, typename std::enable_if<std::is_same_v<T, const char*>>::type* = nullptr)
			: XLOPER12{ .val = {.str = utf8::mbstowcs(str)}, .xltype = xltypeStr }
		{ }
		OPER& operator&=(const XLOPER12& o)
		{
			ensure(type(*this) == xltypeStr && type(o) == xltypeStr);

			XLOPER12 res;
			int ret = Excel12(xlfConcatenate, &res, 2, this, &o);
			ensure(ret == xlretSuccess);
			ensure(type(res) == xltypeStr);
			res.xltype |= xlbitXLFree;
			dealloc();
			if (o.xltype & xlbitXLFree) {
				Excel12(xlFree, 0, 1, &o);
			}
			alloc(res.val.str + 1, res.val.str[0]);

			return *this;
		}

		// Bool
		constexpr explicit OPER(bool xbool)
			: XLOPER12{ Bool(xbool) }
		{ }
		bool operator==(bool xbool) const
		{
			return type(*this) == xltypeBool && (bool)val.xbool == xbool;
		}
		OPER& operator=(bool xbool) noexcept
		{
			dealloc();
			*this = Bool(xbool);

			return *this;
		}

		// Int
		constexpr explicit OPER(int w) noexcept
			: XLOPER12{ Int(w) }
		{ }
		bool operator==(int w) const
		{
			return type(*this) == xltypeInt && val.w == w;
		}

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
				alloc(rows(o), columns(o), Multi(o));
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
				return static_cast<OPER&>(Multi(*this)[i]);
			}
			else {
				ensure(i == 0);
				return *this;
			}
		}
		const OPER& operator[](int i) const
		{
			if (type(*this) == xltypeMulti) {
				return static_cast<OPER&>(Multi(*this)[i]);
			}
			else {
				ensure(i == 0);
				return *this;
			}
		}

		OPER& operator()(int i, int j)
		{
			if (type(*this) == xltypeMulti) {
				return static_cast<OPER&>(Multi(*this)[i * columns(*this) + j]);
			}
			else {
				ensure(i == 0 && j == 0);
				return *this;
			}
		}
		const OPER& operator()(int i, int j) const
		{
			if (type(*this) == xltypeMulti) {
				return static_cast<OPER&>(Multi(*this)[i * columns(*this) + j]);
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
			//ensure(is_object());
			int i = index(str);
			//ensure(i < columns(*this));

			return operator()(1, i);
		}
		const OPER& operator[](const XCHAR* str) const
		{
			//ensure(is_object());
			int i = index(str);
			//ensure(i < columns(*this));

			return operator()(1, i);
		}

	private:
		void dealloc()
		{
			// xltype & xlbitDLLFree is freed when xlAutoFree12 is called.
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
				delete[] static_cast<OPER*>(val.array.lparray);
			}

			xltype = xltypeNil;
		}

		// Str
		constexpr void alloc(const XCHAR* str, XCHAR len)
		{
			xltype = xltypeStr;
			val.str = new XCHAR[1 + static_cast<size_t>(len)];
			val.str[0] = len;
			std::copy_n(str, len, val.str + 1);
		}
		// Multi
		constexpr void alloc(int r, int c, const XLOPER12* a)
		{
			xltype = xltypeMulti;
			val.array.rows = r;
			val.array.columns = c;
			val.array.lparray = nullptr;
			if (r && c) {
				val.array.lparray = new OPER[static_cast<size_t>(r) * c];
				if (a) {
					for (int i = 0; i < r * c; ++i) {
						new (val.array.lparray + i) OPER(a[i]);
					}
				}
			}
		}
	};

	inline OPER Evaluate(const XLOPER12& x)
	{
		XLOPER12 o;

		Excel12(xlfEvaluate, &o, 1, &x);

		return OPER(o);
	}
	// Convert to string using Excel TEXT().
	inline OPER Text(const XLOPER12& x, const OPER& format = OPER(L"General", 7))
	{
		XLOPER12 text;

		Excel12(xlfText, &text, 2, &x, &format);

		return OPER(text);
	}
	// Convert string to value type using VALUE().
	inline OPER Value(const XLOPER12& x)
	{
		XLOPER12 value;

		Excel12(xlfValue, &value, 1, &x);

		return OPER(value);
	}

} // namespace xll

using LPOPER = xll::OPER*;

// Concatenate strings
inline xll::OPER operator&(const XLOPER12& x, const XLOPER12& y)
{
	return std::move(xll::OPER(x) &= y);
}
