// xlerror.h - Error handling
#pragma once
#include <WinUser.h>

enum xll_alert_type {
	XLL_ALERT_ERROR = 1,
	XLL_ALERT_WARNING = 2,
	XLL_ALERT_INFORMATION = 4,
};

extern int xll_alert_level;

inline int XLL_ALERT(LPCSTR text, LPCSTR caption, UINT type = 0)
{
	return MessageBoxA(0, text, caption, MB_OKCANCEL | type);
}
inline int XLL_ERROR(LPCSTR text)
{
	if (xll_alert_level & XLL_ALERT_ERROR) {
		int ret = XLL_ALERT(text, "Error", MB_ICONERROR);
		if (ret == IDCANCEL) {
			xll_alert_level &= ~XLL_ALERT_ERROR;
		}
	}

	return xll_alert_level;
}
inline int XLL_WARNING(LPCSTR text)
{
	if (xll_alert_level & XLL_ALERT_WARNING) {
		int ret = XLL_ALERT(text, "Warning", MB_ICONWARNING);
		if (ret == IDCANCEL) {
			xll_alert_level &= ~XLL_ALERT_WARNING;
		}
	}

	return xll_alert_level;
}
inline int XLL_INFORMATION(LPCSTR text)
{
	if (xll_alert_level & XLL_ALERT_INFORMATION) {
		int ret = XLL_ALERT(text, "Information", MB_ICONINFORMATION);
		if (ret == IDCANCEL) {
			xll_alert_level &= ~XLL_ALERT_INFORMATION;
		}
	}

	return xll_alert_level;
}
