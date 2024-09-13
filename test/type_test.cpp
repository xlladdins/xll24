// type_test.cpp - Test Type
#include "xll.h"
#include "macrofun.h"
#include "type.h"

using namespace xll;

AddIn xai_type(Macro(L"xll_type", L"XLL.TYPE"));
int WINAPI xll_type()
{
#pragma XLLEXPORT
	try {
		Select sel;
		Type("some text", 100, 1000);
		sel.Down();
		Type("more text", 100, 1000);
		sel.Down();
		Type("=len(\"hello\")", 100, 1000);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;
}