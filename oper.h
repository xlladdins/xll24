// oper.h - OPER class
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#ifdef _DEBUG
#include <cassert>
#endif // _DEBUG
#include <initializer_list>
#include "xloper.h"
#include "utf8.h"

namespace xll {

#define XLL_IS_CHAR(S,T) std::is_same<S, typename std::remove_cv<T>::type>::value
	template<class T>
	struct is_char
		: std::integral_constant<bool, XLL_IS_CHAR(char,T) || XLL_IS_CHAR(XCHAR, T)>
	{ };
#undef XLL_IS_CHAR

	// Length of null terminated wide string.
	constexpr XCHAR len(const XCHAR* s, XCHAR n = 0)
	{
		return s && *s && n < 0x4FFF ? len(s + 1, n + 1) : n;
	}
	static_assert(len(L"abc") == 3);

	// Length of null terminated string.
	constexpr char len(const char* s, char n = 0)
	{
		return s && *s && n < 0xFF ? len(s + 1, n + 1) : n;
	}
	static_assert(len("abc") == 3);

	struct OPER : public XLOPER12 {
		// Nil is the default.
		constexpr OPER() noexcept
			: XLOPER12{ Nil }
		{ }

		constexpr OPER(const XLOPER12& o)
			: XLOPER12{ o }
		{
			int otype = type(o);

			if (otype == xltypeStr) {
				alloc(Str(*this), val.str[0]);
			}
			else if (otype == xltypeMulti) {
				alloc(rows(o), columns(o), Multi(o));
			}
			else if (otype == xltypeBigData) {
				alloc(BigData(*this), val.bigdata.cbData);
			}
		}
		constexpr OPER(const OPER& o)
			: OPER(static_cast<const XLOPER12&>(o))
		{ }

		OPER& operator=(const XLOPER12& x)
		{
			if (this != &x) {
				*this = OPER(x); 
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
				dealloc();
				val = o.val;
				xltype = std::exchange(o.xltype, xltypeNil);
			}

			return *this;
		}
		~OPER()
		{
			dealloc();
		}

		// Num - 64-bit IEEE floating point
		constexpr explicit OPER(double num) noexcept
			: XLOPER12{ Num(num) }
		{ }
		OPER& operator=(double num)
		{
			dealloc();

			return *this = Num(num);
		}
		constexpr bool operator==(double num) const noexcept
		{
			return type(*this) == xltypeNum && asNum(*this) == num;
		}

		// Str - Counted wide character string.
		constexpr OPER(const XCHAR* str, XCHAR len)
		{
			alloc(str, len);
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
		// utf8::mbstowcs calls new wchar_t[].
		explicit OPER(const char* str)
			: XLOPER12{ .val = {.str = utf8::mbstowcs(str)}, .xltype = xltypeStr }
		{ }
		constexpr OPER(const std::wstring_view& str)
			: OPER(str.data(), static_cast<XCHAR>(str.size()))
		{ }
		OPER(const std::string_view& str)
			: OPER(str.data())
		{ }
		
		OPER& operator=(const XCHAR* str)
		{
			dealloc();

			return *this = OPER(str);
		}
		OPER& operator=(const char* str)
		{
			dealloc();
			
			return *this = OPER(str);
		}
		OPER& operator=(const std::wstring_view& str)
		{
			dealloc();

			return *this = OPER(str.data(), static_cast<XCHAR>(str.size()));
		}
		OPER& operator=(const std::string_view& str)
		{
			dealloc();
			
			return *this = OPER(str.data());
		}
		constexpr bool operator==(const XCHAR* str) const
		{
			return type(*this) == xltypeStr && view(*this) == str;
		}
		constexpr bool operator==(const char* str) const
		{
			if (type(*this) != xltypeStr) {
				return false;
			}
			
			int i = 0;
			while (str[i] && i < val.str[0]) {
				if (val.str[i + 1] != str[i]) {
					return false;
				}
				++i;
			}
			
			return str[i] == 0;
		}

		OPER& operator&=(const XLOPER12& o)
		{
			if (size(*this) == 0) {
				operator=(o);
			}
			else {
				ensure(type(*this) == xltypeStr);
				ensure(type(o) == xltypeStr);

				XCHAR len = val.str[0];
				XCHAR olen = o.val.str[0];
				OPER res = OPER(nullptr, len + olen);
				//res.val.str[0] = len + olen;
				std::copy_n(val.str + 1, len, res.val.str + 1);
				std::copy_n(o.val.str + 1, olen, res.val.str + 1 + len);
				std::swap(val.str, res.val.str);
			}

			return *this;
		}
		/*
		OPER& operator&=(const XCHAR* str)
		{
			return operator&=(OPER(str));
		}
		OPER& operator&=(const char* str)
		{
			return operator&=(OPER(str));
		}
		OPER& operator&=(const std::wstring_view& str)
		{
			return operator&=(OPER(str));
		}
		OPER& operator&=(const char* str)
		{
			return operator&=(OPER(str));
		}
		*/

		// Bool
		constexpr explicit OPER(bool xbool)
			: XLOPER12{ Bool(xbool) }
		{ }
		bool operator==(bool xbool) const
		{
			return type(*this) == xltypeBool && Bool(*this) == static_cast<BOOL>(xbool);
		}
		OPER& operator=(bool xbool) noexcept
		{
			dealloc();
			
			return *this = Bool(xbool);
		}

		// SRef
		constexpr explicit OPER(const XLREF12& ref) noexcept
			: XLOPER12{ SRef(ref) }
		{ }
		bool operator==(const XLREF12& ref) const
		{
			return type(*this) == xltypeSRef && val.sref.ref == ref;
		}
		OPER& operator=(const XLREF12& ref) noexcept
		{
			dealloc();

			return *this = SRef(ref);
		}

		// Int
		constexpr explicit OPER(int w) noexcept
			: XLOPER12{ Int(w) }
		{ }
		bool operator==(int w) const
		{
			switch (type(*this)) {
				case xltypeInt:
					return val.w == w;
				case xltypeNum:
					return val.num == w;
				case xltypeBool:
					return val.xbool == w;
			}

			return false;
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
				xltype &= ~xlbitXLFree;
				::Excel12v(xlFree, 0, 1, (LPXLOPER12*)this);
			}
			else if (xltype == xltypeStr) {
				delete[] val.str;
			}
			else if (xltype == xltypeMulti) {
				for (auto& o : span(*this)) {
					static_cast<OPER&>(o).dealloc();
				}
				delete[] static_cast<OPER*>(Multi(*this));
			}
			else if (xltype == xltypeBigData) {
				delete[] BigData(*this);
			}

			xltype = xltypeNil;
		}

		// Str
		constexpr void alloc(const XCHAR* str, XCHAR len)
		{
			xltype = xltypeStr;
			val.str = new XCHAR[1 + static_cast<size_t>(len)];
			val.str[0] = len;
			if (str && len) {
				std::copy_n(str, len, val.str + 1);
			}
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
			else {
				xltype = xltypeNil;
			}
		}
		// BigData
		constexpr void alloc(const BYTE* data, int len)
		{
			xltype = xltypeBigData;
			val.bigdata.cbData = len;
			val.bigdata.h.lpbData = new BYTE[len];
			std::copy_n(data, len, val.bigdata.h.lpbData);
		}
	};

} // namespace xll

using LPOPER = xll::OPER*;

static_assert(sizeof(xll::OPER) == sizeof(XLOPER12));

// String concatenation like Excel.
inline xll::OPER operator&(const XLOPER12& x, const XLOPER12& y)
{
	return xll::OPER(x) &= y;
}
