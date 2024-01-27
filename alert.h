// alert.h - Error, Warning, and Information alerts
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
// Store alert mask in registry to persist across sessions.
#pragma once
#include <string_view>
#include <windows.h>
#include "XLCALL.H"

enum xll_alert_level {
	XLL_ALERT_ERROR = 1,
	XLL_ALERT_WARNING = 2,
	XLL_ALERT_INFORMATION = 4,
};

#define XLL_SUB_KEY "Software\\KALX\\xll"
#define XLL_VALUE_NAME "AlertMask"

inline int get_alert_mask() noexcept
{
	HKEY hkey{ 0 };
	DWORD disp{ 0 };
	DWORD data = XLL_ALERT_ERROR | XLL_ALERT_WARNING | XLL_ALERT_INFORMATION;

	LSTATUS status = RegCreateKeyExA(HKEY_CURRENT_USER, XLL_SUB_KEY, 0, nullptr, 0, KEY_READ, 0, &hkey, &disp);
	if (status == ERROR_SUCCESS) {
		DWORD type{ 0 };
		DWORD size = sizeof(data);
		status = RegQueryValueExA(hkey, XLL_VALUE_NAME, 0, &type, (LPBYTE)&data, &size);
		if (status == ERROR_FILE_NOT_FOUND) {
			RegSetValueExA(hkey, XLL_VALUE_NAME, 0, REG_DWORD, (LPBYTE)&data, sizeof(DWORD));
		}
	}

	return data;
}
inline void set_alert_mask(int level)
{
	HKEY hkey;
	DWORD disp;

	LSTATUS status = RegCreateKeyExA(HKEY_CURRENT_USER, XLL_SUB_KEY, 0, 0, 0, KEY_WRITE, 0, &hkey, &disp);
	if (status == ERROR_SUCCESS) {
		status = RegSetValueExA(hkey, XLL_VALUE_NAME, 0, REG_DWORD, (LPBYTE)&level, sizeof(DWORD));
	}
}

// Handle to Excel window.
inline HWND xllGetHwnd(void) noexcept
{
	XLOPER12 xHwnd = { .xltype = xltypeNil };

	const int ret = Excel12(xlGetHwnd, &xHwnd, 0);
	if (ret != xlretSuccess || xHwnd.xltype != xltypeInt) {
		return NULL;
	}

	return static_cast<HWND>(IntToPtr(xHwnd.val.w));
}

inline int XLL_ALERT(int level, std::string_view text, LPCSTR caption, UINT type = 0)
{
	int alert_level = get_alert_mask();

	if (alert_level & level) {
		const int ret = MessageBoxA(xllGetHwnd(), std::string(text).c_str(), caption, MB_OKCANCEL | type);
		if (ret == IDCANCEL) {
			alert_level &= ~level;
			set_alert_mask(alert_level);
		}
	}

	return alert_level;
}
inline int XLL_ALERT(int level, std::wstring_view text, LPCWSTR caption, UINT type = 0)
{
	int alert_level = get_alert_mask();

	if (alert_level & level) {
		int ret = MessageBoxW(xllGetHwnd(), std::wstring(text).c_str(), caption, MB_OKCANCEL | type);
		if (ret == IDCANCEL) {
			alert_level &= ~level;
			set_alert_mask(alert_level);
		}
	}

	return alert_level;
}

inline int XLL_ERROR(std::string_view text)
{
	return XLL_ALERT(XLL_ALERT_ERROR, text, "Error", MB_ICONERROR);
}
inline int XLL_ERROR(std::wstring_view text)
{
	return XLL_ALERT(XLL_ALERT_ERROR, text, L"Error", MB_ICONERROR);
}

inline int XLL_WARNING(std::string_view text)
{
	return XLL_ALERT(XLL_ALERT_WARNING, text, "Warning", MB_ICONWARNING);
}
inline int XLL_WARNING(std::wstring_view text)
{
	return XLL_ALERT(XLL_ALERT_WARNING, text, L"Warning", MB_ICONWARNING);
}

inline int XLL_INFORMATION(std::string_view text)
{
	return XLL_ALERT(XLL_ALERT_INFORMATION, text, "Information", MB_ICONINFORMATION);
}
inline int XLL_INFORMATION(std::wstring_view text)
{
	return XLL_ALERT(XLL_ALERT_INFORMATION, text, L"Information", MB_ICONINFORMATION);
}
