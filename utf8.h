#pragma once
#include <stringapiset.h>
#include <string_view>
#include <memory>

namespace utf8 {

	//!!! return std::wstring_view to allocated string
	
	// Multi-byte character string to counted wide character string allocated by new[].
	// Returned string is null terminated if n is -1.
	inline wchar_t* mbstowcs(const char* s, int n = -1)
	{
		wchar_t* ws = nullptr;

		if (!s || n == 0) { // empty string
			ws = new wchar_t[1];
			ws[0] = 0;

			return ws;
		}

		int wn = MultiByteToWideChar(CP_UTF8, 0, s, n, nullptr, 0);
		if (wn == 0) {
			return nullptr;
		}

		ws = new wchar_t[static_cast<size_t>(wn) + 2];
		if (ws) {
			if (wn != MultiByteToWideChar(CP_UTF8, 0, s, n, ws + 1, wn)) {
				delete [] ws;

				return nullptr;
			}
			ensure(wn <= WCHAR_MAX/2);
			// MBTWC includes terminating null if n == -1
			ws[0] = static_cast<wchar_t>(wn - (n == -1));
			ws[ws[0] + 1] = 0;
		}

		return ws;
	}
	inline std::wstring mbstowstring(const char* s, int n = -1)
	{
		std::wstring ws;

		wchar_t* pws = mbstowcs(s, n);
		ensure(pws);
		ws.assign(pws + 1, pws[0]);
		delete [] pws;

		return ws;
	}

#ifdef _DEBUG
	inline int test()
	{
		using uptr = std::unique_ptr<wchar_t[]>;
		{
			uptr s{mbstowcs(nullptr)};
			ensure(s[0] == 0);
		}
		{
			uptr s{ mbstowcs("")};
			ensure(s[0] == 0);
		}
		{
			uptr s{ mbstowcs("abc", 0) };
			ensure(s[0] == 0);
		}
		{
			uptr s{utf8::mbstowcs("")};
			ensure(s[0] == 0);
			ensure(s[1] == L'\0'); // null terminated
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
			auto ws = utf8::mbstowstring("abc");
			ensure(ws[0] == L'a');
			ensure(ws[1] == L'b');
			ensure(ws[2] == L'c');
			ensure(ws[3] == L'\0'); // null terminated
		}
		return 0;
	}
#endif // _DEBUG

} // namespace utf8
