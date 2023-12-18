// xlauto.cpp - xlAutoXXX functions
#include <stdexcept>
#include "xll.h"

int xll_alert_level = XLL_ALERT_ERROR | XLL_ALERT_WARNING | XLL_ALERT_INFORMATION;

// Called by Excel when the xll is opened.
extern "C" int __declspec(dllexport) WINAPI
xlAutoOpen(void)
{
	try {
		XLL_INFORMATION("info");
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");

		return FALSE;
	}

	return TRUE;
}