// oper.h - OPER class
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#ifdef _DEBUG
#include <cassert>
#endif // _DEBUG
#include <initializer_list>
#include <map>
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

	struct OPER;
	// Keep track of nested handles.
	struct handles {
		static double handle(OPER* po)
		{
			return static_cast<double>(reinterpret_cast<uintptr_t>(po));
		}
		static OPER* pointer(double h)
		{
			return reinterpret_cast<OPER*>(static_cast<uintptr_t>(h));
		}
		static std::map<double, OPER*>& map()
		{
			static std::map<double, OPER*> map;

			return map;
		}
		static double insert(OPER* po)
		{
			map().insert({ handle(po), po});

			return handle(po);
		}
		static OPER* find(double h)
		{
			const auto i = map().find(h);

			return i == map().end() ? nullptr : i->second;
		}
		static void erase(double h)
		{
			map().erase(h);
		}
	};

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

		const OPER* begin() const
		{
			return &operator[](0);
		}
		OPER* begin()
		{
			return &operator[](0);
		}
		const OPER* end() const
		{
			return begin() + (isMulti(*this) ? size(*this) : 1);
		}
		OPER* end()
		{
			return begin() + (isMulti(*this) ? size(*this) : 1);
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
		std::wstring to_wstring() const
		{
			ensure(isStr(*this));

			return std::wstring(val.str + 1, val.str[0]);
		}
		std::string to_string() const
		{
			ensure(isStr(*this));

			const char* s = utf8::wcstombs(val.str + 1, val.str[0]);
			std::string str(s + 1, s[0]);
			delete[] s;

			return str;
		}

		// Replace non-alphanumeric characters with '_'.
		OPER safe_name() const
		{
			if (isStr(*this)) {
				OPER safe(*this);
				for (int i = 1; i <= safe.val.str[0]; ++i) {
					if (!iswalnum(safe.val.str[i])) {
						safe.val.str[i] = L'_';
					}
				}

				return safe;
			}

			return ErrValue;
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
			return isBool(*this) && Bool(*this) == static_cast<BOOL>(xbool);
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
			return isSRef(*this) && val.sref.ref == ref;
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
		// Excel converts ints to doubles.
		template<std::integral T>
		bool operator==(T w) const
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
			ensure(r * c == size(*this));

			val.array.rows = r;
			val.array.columns = c;

			return *this;
		}
		OPER& resize(int r, int c)
		{
			if (r * c == size(*this)) {
				return reshape(r, c);
			}

			OPER o(*this);
			dealloc();
			alloc(r, c, nullptr);
			for (int i = 0; i < (std::min)(r * c, size(o)); ++i) {
				operator[](i) = o[i];
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
			if (isMulti(*this)) {
				return static_cast<OPER&>(Multi(*this)[i * columns(*this) + j]);
			}
			else {
				//ensure(i == 0 && j == 0);
				return *this;
			}
		}
		const OPER& operator()(int i, int j) const
		{
			if (isMulti(*this)) {
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
		OPER& operator[](const std::string_view& key)
		{
			return static_cast<OPER&>(value(*this, OPER(key)));
		}
		OPER& operator[](const std::wstring_view& key)
		{
			return static_cast<OPER&>(value(*this, OPER(key)));
		}

		const OPER& operator[](const OPER& key) const
		{
			return value(*this, key);
		}
		const OPER& operator[](const std::string_view& key) const
		{
			return value(*this, OPER(key));
		}
		const OPER& operator[](const std::wstring_view& key) const
		{
			return value(*this, OPER(key));
		}

		// Promote to 1 x 1 multi.
		OPER& enlist()
		{
			if (!isMulti(*this)) {
				OPER o(*this);
				dealloc();
				alloc(1, 1, &o);
			}

			return *this;
		}

		OPER& vstack(const XLOPER12& x)
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
			if (!isMulti(*this)) {
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
			if (isMulti(*this)) {
				XLOPER12 o;
				int ret = ::Excel12(xlfTranspose, &o, 1, this);
				if (ret == xlretSuccess) {
					operator=(o);
					::Excel12(xlFree, 0, 1, &o);
				}
			}

			return *this;
		}
		OPER& hstack(const XLOPER12& x)
		{
			// Not efficient.
			OPER x_(*this);
			x_.transpose();
			OPER _x(x);
			_x.transpose();
			vstack(_x);
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
				if (count(*this)) {
					delete[] BigData(*this);
				}
			}
			else if (xltype == xltypeNum) {
				OPER* po = handles::find(val.num);
				if (po) {
					po->dealloc();
				}
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
			if (len) {
				val.bigdata.h.lpbData = new BYTE[len];
				std::copy_n(data, len, val.bigdata.h.lpbData);
			}
		}
		constexpr void alloc(const XLOPER12& x)
		{
			xltype = type(x);
			switch (xltype) {
			case xltypeStr:
				alloc(Str(x), count(x));
				break;
			case xltypeMulti:
				alloc(rows(x), columns(x), Multi(x));
				break;
			case xltypeBigData:
				if (count(x)) {
					alloc(BigData(x), count(x));
				}
				else { // handle
					val.bigdata.h.hdata = x.val.bigdata.h.hdata;
				}
				break;
			default:
				val = x.val;
			}

		}
	};

	// Replace nested OPER with safe handles.
	inline OPER compress(const OPER& o)
	{
		if (type(o) != xltypeMulti) {
			return o;
		}

		OPER o_(o);
		for (OPER& oi : o_) {
			if (isMulti(oi)) {
				oi = handles::insert(new OPER(compress(oi)));
			}
		}

		return o_;
	}

	inline OPER expand(const OPER& o)
	{
		if (isMulti(o)) {
			OPER o_(o);
			for (OPER& oi : o_) {
				oi = expand(oi);
			}

			return o_;
		}
		if (type(o) == xltypeNum) {
			OPER* po = handles::find(o.val.num);
			if (po) {
				return expand(*po);
			}
		}

		return o;


	}

} // namespace xll

using LPOPER = xll::OPER*;

static_assert(sizeof(xll::OPER) == sizeof(XLOPER12));

// String concatenation like Excel.
inline xll::OPER operator&(const XLOPER12& x, const XLOPER12& y)
{
	return xll::OPER(x) &= y;
}
