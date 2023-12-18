#pragma once
#ifdef _DEBUG
#include <cassert>
#endif // _DEBUG
#include <algorithm>
#include <compare>
#include <string_view>
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include "XLCALL.H"

namespace xll {

	constexpr XLOPER12 Num12(double num)
	{
		return XLOPER12{ .val = {.num = num}, .xltype = xltypeNum };
	}
	static_assert(Num12(1.23).val.num == 1.23);
	static_assert(Num12(1.23).xltype == xltypeNum);

	struct Str12 : public XLOPER12 {
		// NULL string
		constexpr Str12()
			: XLOPER12{ .val = {.str = nullptr}, .xltype = xltypeStr }
		{ }
		constexpr Str12(const XCHAR* str, XCHAR len)
		{
			xltype = xltypeStr;
			val.str = nullptr;
			if (str) {
				val.str = const_cast<wchar_t*>(new wchar_t[1 + len]);
				val.str[0] = len;
				std::copy_n(str, len, val.str + 1);
			}
		}
		constexpr Str12(const XCHAR* str)
			: Str12(str + 1, str ? str[0] : 0)
		{ }
		constexpr Str12(const Str12& str)
			: Str12(str.val.str)
		{ }
		Str12& operator=(const Str12& str)
		{
			delete[] val.str;

			new (this) Str12(str);
		}
		constexpr ~Str12()
		{
			delete[] val.str;
		}

		constexpr auto operator<=>(const Str12& str) const
		{
			if (val.str && str.val.str) {
				return std::lexicographical_compare_three_way(val.str, val.str + 1 + val.str[0], str.val.str, str.val.str + 1 + str.val.str[0]);
			}
			else if (val.str) {
				return 1 <=> 0;
			}
			else if (str.val.str) {
				return -1 <=> 0;
			}
			else {
				return 0 <=> 0;
			}
		}

		constexpr std::wstring_view view() const
		{
			return std::wstring_view(val.str + 1, val.str[0]);
		}

#ifdef _DEBUG
		static int test()
		{
			{
				Str12 s;
			}

			return 0;
		}
#endif // _DEBUG
	};
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
	struct OPER12 : public XLOPER12 {
		constexpr OPER12() noexcept
			: XLOPER12{ .xltype = xltypeNil }
		{ }
		constexpr explicit OPER12(double num) noexcept
			: XLOPER12{ Num12(num) }
		{ }
		constexpr explicit OPER12(const XCHAR* str)
		{
			new (this) Str12(str, static_cast<XCHAR>(wcslen(str)));
		}
		constexpr explicit OPER12(int w) noexcept
			: XLOPER12{ .val = {.w = w}, .xltype = xltypeInt }
		{ }
		constexpr OPER12(const OPER12& o)
			: XLOPER12{ o }
		{
			if (xltype == xltypeStr) {
				new (this) Str12(o.val.str + 1, o.val.str[0]);
			}
		}
		OPER12& operator=(const OPER12& o)
		{
			if (xltype == xltypeStr) {
				delete[] val.str;
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
			if (xltype & xlbitXLFree) {
				Excel12v(xlFree, 0, 1, (LPXLOPER12*)this);
			} 
			else if (xltype == xltypeStr) {
				delete[] val.str;
			}
		}

		const OPER12& operator[](int i) const
		{
			if (xltype == xltypeMulti) {
				return (const OPER12&)val.array.lparray[i];
			}
			else {
				// ensure i == 0
				return *this;
			}
		}
	};

} // namespace xll