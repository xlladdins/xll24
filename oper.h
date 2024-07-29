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
		// xltypeNil
		constexpr OPER() noexcept
			: XLOPER12{ Nil }
		{ }

		constexpr OPER(const XLOPER12& x)
			: XLOPER12{x}
		{
			if (isAlloc(x)) {
				alloc(x);
			}
		}
		constexpr OPER(const OPER& o)
			: OPER(static_cast<const XLOPER12&>(o))
		{ }

		OPER& operator=(const XLOPER12& x) noexcept
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
			: XLOPER12{ o }
		{
			o.xltype = xltypeNil;
		}
		OPER& operator=(OPER&& o) noexcept
		{
			if (this != &o) {
				dealloc();
				xltype = std::exchange(o.xltype, xltypeNil);
				val = o.val;
			}

			return *this;
		}
		~OPER()
		{
			dealloc();
		}

		constexpr explicit operator bool() const noexcept
		{
			return isTrue(*this);
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
			return asNum(*this) == num;
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
		
		// Excel string concatenation.
		OPER& operator&=(const XLOPER12& o)
		{
			if (size(*this) == 0) {
				operator=(o);
			}
			else {
				if (type(*this) != xltypeStr || type(o) != xltypeStr) {
					dealloc();
					*this = ErrNA;
				}
				else {
					XCHAR len = val.str[0];
					XCHAR olen = o.val.str[0];
					OPER res(nullptr, len + olen);
					//res.val.str[0] = len + olen;
					std::copy_n(val.str + 1, len, res.val.str + 1);
					std::copy_n(o.val.str + 1, olen, res.val.str + 1 + len);
					dealloc();
					operator=(std::move(res));
				}
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
		OPER& operator=(int w) noexcept
		{
			dealloc();

			return *this = Int(w);
		}
		bool operator==(int w) const
		{
			return operator==(static_cast<double>(w));
		}


		// Err
		constexpr explicit OPER(xll::xlerr err) noexcept
			: XLOPER12{ Err(static_cast<int>(err)) }
		{ }

		// Multi
		constexpr OPER(int r, int c, const XLOPER12* a = nullptr)
		{
			alloc(r, c, a);
		}
		// One row multi.
		constexpr OPER(std::initializer_list<XLOPER12> a)
			: OPER(1, static_cast<int>(a.size()), a.begin())
		{ }

		OPER& reshape(int r, int c)
		{
			if (r * c != size(*this)) {
				dealloc();
				xltype = xltypeErr;
				val.err = xlerrValue;
			}
			else {
				val.array.rows = r;
				val.array.columns = c;
			}

			return *this;
		}

		// One-dimensional index.
		OPER& operator[](int i)
		{
			if (type(*this) == xltypeMulti) {
				return static_cast<OPER&>(Multi(*this)[i]);
			}
			else {
				//ensure(i == 0);
				return *this;
			}
		}
		const OPER& operator[](int i) const
		{
			if (type(*this) == xltypeMulti) {
				return static_cast<OPER&>(Multi(*this)[i]);
			}
			else {
				//ensure(i == 0);
				return *this;
			}
		}

		// Two-dimensional index.
		OPER& operator()(int i, int j)
		{
			if (type(*this) == xltypeMulti) {
				return static_cast<OPER&>(Multi(*this)[i * columns(*this) + j]);
			}
			else {
				//ensure(i == 0 && j == 0);
				return *this;
			}
		}
		const OPER& operator()(int i, int j) const
		{
			if (type(*this) == xltypeMulti) {
				return static_cast<OPER&>(Multi(*this)[i * columns(*this) + j]);
			}
			else {
				//ensure(i == 0 && j == 0);
				return *this;
			}
		}
		// JSON lookup for 2 row multi.
		OPER& operator[](const OPER& key)
		{
			return static_cast<OPER&>(value(*this, key));
		}
		const OPER& operator[](const OPER& key) const
		{
			return value(*this, key);
		}
		
		// Promote to 1 x 1 multi.
		OPER& enlist()
		{
			if (type(*this) != xltypeMulti) {
				OPER o(*this);
				dealloc();
				alloc(1, 1, &o);
			}

			return *this;
		}

		OPER& push_bottom(const XLOPER12& x)
		{
			if (size(x) == 0) {
				return *this;
			}
			if (size(*this) == 0) {
				return operator=(x);
			}
			if (columns(*this) != columns(x)) {
				return operator=(ErrValue);
			}
			if (!isMulti(*this)) {
				enlist();
			}

			int n = size(*this);
			OPER o(rows(*this) + rows(x), columns(*this), nullptr);
			for (int i = 0; i < n; ++i) {
				o[i] = operator[](i);
			}
			if (type(x) != xltypeMulti) {
				o[n] = x;
			}
			else {
				for (int i = 0; i < size(x); ++i) {
					o[size(*this) + i] = x.val.array.lparray[i];
				}
			}
			std::swap(*this, o);

			return *this;
		}
		OPER& transpose()
		{
			xll::transpose(*this);

			return *this;
		}
		OPER& push_right(const XLOPER12& x)
		{
			// Not efficient.
			OPER x_(*this);
			x_.transpose();
			OPER _x(x);
			_x.transpose();
			push_bottom(_x);
			x_.transpose();
			std::swap(*this, x_);

			return *this;
		}

		// Append single item to row or column vector
		OPER& append(const XLOPER12& x)
		{
			if (size(*this) == 0) {
				operator=(x);

				return enlist();
			}
			if (!isMulti(*this)) {
				enlist();
			}
			if (rows(*this) != 1 && columns(*this) != 1) {
				return operator=(ErrValue);
			}
			int n = size(*this);
			OPER o(1, n + 1);
			for (int i = 0; i < n; ++i) {
				o[i] = operator[](i);
			}
			o[n] = x;

			std::swap(*this, o);	
			if (rows(*this) != 1) {
				std::swap(val.array.rows, val.array.columns);
			}

			return *this;
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
			if (!val.str) {
				xltype = xltypeErr;
				val.err = xlerrNA;
			}
			else {
				val.str[0] = len;
				if (str && len) {
					std::copy_n(str, len, val.str + 1);
				}
			}
		}
		// Multi
		constexpr void alloc(int r, int c, const XLOPER12* a)
		{
			xltype = xltypeMulti;
			val.array.rows = r;
			val.array.columns = c;
			val.array.lparray = nullptr;
			if (size(*this)) {
				val.array.lparray = new OPER[size(*this)];
				if (a) {
					for (int i = 0; i < r * c; ++i) {
						new (val.array.lparray + i) OPER(a[i]);
					}
				}
			}
			else {
				xltype = xltypeErr;
				val.err = xlerrNA;
			}
		}
		// BigData
		constexpr void alloc(const BYTE* data, long len)
		{
			xltype = xltypeBigData;
			val.bigdata.cbData = len;
			val.bigdata.h.lpbData = new BYTE[len];
			std::copy_n(data, len, val.bigdata.h.lpbData);
		}
		constexpr void alloc(const XLOPER12& x)
		{
			xltype = type(x);
			switch (xltype) {
			case xltypeStr:
				alloc(Str(x), x.val.str[0]);
				break;
			case xltypeMulti:
				alloc(x.val.array.rows, x.val.array.columns, Multi(x));
				break;
			case xltypeBigData:
				alloc(BigData(x), x.val.bigdata.cbData);
				break;
			default:
				val = x.val;
			}

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
