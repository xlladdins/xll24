// utf8.h - utf8 to wide character string conversion
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <stringapiset.h>
#include <string_view>
#include <memory>

namespace utf8 {

	// Wide character string size of multi-byte character string.
	// Default to null terminated string.
	inline int wcslen(const char* s, int n = -1)
	{
		return MultiByteToWideChar(CP_UTF8, 0, s, n, nullptr, 0);
	}
	// Fill ws with wide character string from multi-byte character string.
	inline int mbstowcs(const char* s, int n, wchar_t* ws, int wn)
	{
		return MultiByteToWideChar(CP_UTF8, 0, s, n, ws, wn);
	}

	// Multi-byte character string to counted wide character string allocated by new[].
	// Returned string is null terminated if n is -1 or nullptr on failure.
	inline wchar_t* mbstowcs(const char* s, int n = -1)
	{
		wchar_t* ws = nullptr;

		if (!s || n == 0) { // empty string
			ws = new wchar_t[1];
			ws[0] = 0;

			return ws;
		}

		const int wn = wcslen(s, n);
		if (wn == 0) {
			return nullptr;
		}

		ws = new wchar_t[static_cast<size_t>(wn) + 2];
		if (ws) {
			if (wn != mbstowcs(s, n, ws + 1, wn)) {
				delete [] ws;

				return nullptr;
			}
			if (wn > WCHAR_MAX / 2) {
				delete[] ws;

				return nullptr;
			}
			ws[0] = static_cast<wchar_t>(wn - (n == -1));
		}

		return ws;
	}
	
	inline std::wstring mbstowstring(const char* s, int n = -1)
	{
		std::wstring ws;

		if (s && n != 0) {
			const int wn = wcslen(s, n);
			if (wn != 0) {
				ws.resize(wn);
				if (wn == mbstowcs(s, n, ws.data(), wn)) {
					ws.resize(static_cast<size_t>(wn) - (n == -1)); // remove null terminator
				}
			}
		}

		return ws;
	}

	// Multi-byte character string size of wide character string.
	// Default to null terminated string.
	inline int mbslen(const wchar_t* ws, int wn = -1)
	{
		return WideCharToMultiByte(CP_UTF8, 0, ws, wn, nullptr, 0, 0, 0);
	}
	// Fill ws with wide character string from multi-byte character string.
	inline int wcstombs(const wchar_t* ws, int wn, char* s, int n)
	{
		return WideCharToMultiByte(CP_UTF8, 0, ws, wn, s, n, 0, 0);
	}


	// Multi-byte character string to counted wide character string allocated by new[].
	// Returned string is null terminated if n is -1 or nullptr on failure.
	inline char* wcstombs(const wchar_t* ws, int wn = -1)
	{
		char* s = nullptr;

		if (!ws || wn == 0) { // empty string
			s = new char[1];
			s[0] = 0;

			return s;
		}

		const int n = mbslen(ws, wn);
		if (n == 0) {
			return nullptr;
		}

		s = new char[static_cast<size_t>(n) + 2];
		if (s) {
			if (n != wcstombs(ws, wn, s + 1, n)) {
				delete[] s;

				return nullptr;
			}
			if (n > CHAR_MAX / 2) {
				delete[] s;

				return nullptr;
			}
			s[0] = static_cast<char>(n - (wn == -1));
		}

		return s;
	}

	inline std::string wcstostring(const wchar_t* ws, int wn = -1)
	{
		std::string s;

		if (ws && wn != 0) {
			const int n = mbslen(ws, wn);
			if (n != 0) {
				s.resize(n);
				if (n == wcstombs(ws, wn, s.data(), n)) {
					s.resize(static_cast<size_t>(n) - (wn == -1)); // remove null terminator
				}
			}
		}

		return s;
	}

	//#ifdef _DEBUG
	inline int test()
	{
		{
			using uptr = std::unique_ptr<wchar_t[]>;
			{
				uptr s{ mbstowcs(nullptr) };
				ensure(s[0] == 0);
			}
			{
				uptr s{ mbstowcs("") };
				ensure(s[0] == 0);
			}
			{
				uptr s{ mbstowcs("abc", 0) };
				ensure(s[0] == 0);
			}
			{
				uptr s{ utf8::mbstowcs("") };
				ensure(s[0] == 0);
			}
			{
				uptr s{ utf8::mbstowcs("abc") };
				ensure(s[0] == 3);
				ensure(s[1] == L'a');
				ensure(s[2] == L'b');
				ensure(s[3] == L'c');
				ensure(s[4] == L'\0'); // null terminated
			}
			{
				uptr s{ utf8::mbstowcs("abc", 2) };
				ensure(s[0] == 2);
				ensure(s[1] == L'a');
				ensure(s[2] == L'b');
				// ensure(s[3] == L'\0'); // not null terminated
			}
			{
				ensure(mbstowstring("abc") == L"abc");
				ensure(mbstowstring("abc", 2) == L"ab");
				ensure(mbstowstring("abc", 0) == L"");
				ensure(mbstowstring("") == L"");
				ensure(mbstowstring(nullptr) == L"");
			}
		}
		{
			using uptr = std::unique_ptr<char[]>;
			{
				uptr s{ wcstombs(nullptr) };
				ensure(s[0] == 0);
			}
			{
				uptr s{ wcstombs(L"") };
				ensure(s[0] == 0);
			}
			{
				uptr s{ wcstombs(L"abc", 0) };
				ensure(s[0] == 0);
			}
			{
				uptr s{ utf8::wcstombs(L"") };
				ensure(s[0] == 0);
			}
			{
				uptr s{ utf8::wcstombs(L"abc") };
				ensure(s[0] == 3);
				ensure(s[1] == 'a');
				ensure(s[2] == 'b');
				ensure(s[3] == 'c');
				ensure(s[4] == '\0'); // null terminated
			}
			{
				uptr s{ utf8::wcstombs(L"abc", 2) };
				ensure(s[0] == 2);
				ensure(s[1] == 'a');
				ensure(s[2] == 'b');
				// ensure(s[3] == '\0'); // not null terminated
			}
			{
				ensure(wcstostring(L"abc") == "abc");
				ensure(wcstostring(L"abc", 2) == "ab");
				ensure(wcstostring(L"abc", 0) == "");
				ensure(wcstostring(L"") == "");
				ensure(wcstostring(nullptr) == "");
			}
		}

		return 0;
	}
//#endif // _DEBUG

} // namespace utf8
