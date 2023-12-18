// xlauto.cpp - xlAutoXXX functions
#include <stdexcept>
#include "xll.h"

using namespace xll;

int xll_alert_level = XLL_ALERT_ERROR | XLL_ALERT_WARNING | XLL_ALERT_INFORMATION;

// Called by Excel when the xll is opened.
extern "C" int __declspec(dllexport) WINAPI
xlAutoOpen(void)
{
	try {
		//XLL_ERROR("error");
		//XLL_WARNING("warning");
		//XLL_INFORMATION("info");
		OPER12 o(L"abc");
		auto o2{ o };
		o = o2;
		auto od = Excel12(xlfDate, 2023, 1, 2);
		//auto of = Excel12(xlcFormatNumber, od.val.num, L"yyyy-mm-dd");
		//of = of;
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