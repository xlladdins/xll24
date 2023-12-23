// xlerror.h - Error handling
#pragma once
#include <Windows.h>
#include "XLCALL.H"

enum xll_alert_type {
	XLL_ALERT_ERROR = 1,
	XLL_ALERT_WARNING = 2,
	XLL_ALERT_INFORMATION = 4,
};

inline int get_alert_level()
{
	HKEY hkey;
	DWORD disp;
	DWORD data = 0x7;
	LSTATUS status = RegCreateKeyExA(HKEY_CURRENT_USER, "Software\\KALX\\XLL", 0, 0, 0, KEY_READ, 0, &hkey, &disp);
	if (status == ERROR_SUCCESS) {
		DWORD type, size = sizeof(data);
		status = RegQueryValueExA(hkey, "AlertLevel", 0, &type, (LPBYTE)&data, &size);
		if (status == ERROR_FILE_NOT_FOUND) {
			RegSetValueExA(hkey, "AlertLevel", 0, REG_DWORD, (LPBYTE)&data, sizeof(DWORD));
		}
	}

	return data;
}
inline void set_alert_level(int level)
{
	HKEY hkey;
	DWORD disp;
	LSTATUS status = RegCreateKeyExA(HKEY_CURRENT_USER, "Software\\KALX\\XLL", 0, 0, 0, KEY_WRITE, 0, &hkey, &disp);
	if (status == ERROR_SUCCESS) {
		status = RegSetValueExA(hkey, "AlertLevel", 0, REG_DWORD, (LPBYTE)&level, sizeof(DWORD));
	}
}

inline HWND xllGetHwnd(void)
{
	XLOPER12 xHwnd = { .xltype = xltypeNil };
	int ret = Excel12(xlGetHwnd, &xHwnd, 0);
	if (ret != xlretSuccess || xHwnd.xltype != xltypeInt) {
		return NULL;
	}

	return (HWND)IntToPtr(xHwnd.val.w);
}

inline int XLL_ALERT(int level, LPCSTR text, LPCSTR caption, UINT type = 0)
{
	int alert_level = get_alert_level();
	if (alert_level & level) {
		int ret = MessageBoxA(xllGetHwnd(), text, caption, MB_OKCANCEL | type);
		if (ret == IDCANCEL) {
			alert_level &= ~level;
			set_alert_level(alert_level);
		}
	}

	return alert_level;
}
inline int XLL_ERROR(LPCSTR text)
{
	return XLL_ALERT(XLL_ALERT_ERROR, text, "Error", MB_ICONERROR);
}

inline int XLL_WARNING(LPCSTR text)
{
	return XLL_ALERT(XLL_ALERT_WARNING, text, "Warning", MB_ICONWARNING);
}

inline int XLL_INFORMATION(LPCSTR text)
{
	return XLL_ALERT(XLL_ALERT_INFORMATION, text, "Information", MB_ICONINFORMATION);
}
