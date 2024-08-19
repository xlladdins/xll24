// range.cpp - OPER range handles
// 
#include "xll.h"

using namespace xll;

AddIn xai_range_set(
	Function(XLL_HANDLEX, L"xll_range_set", L"\\RANGE")
	.Arguments({
		Arg(XLL_LPOPER, L"Range", L"is a range.", "={1,2,3;4,5,6}")
		})
	.Uncalced()
	.Category(L"XLL")
	.FunctionHelp(L"Return a handle to a range.")
);
HANDLEX WINAPI xll_range_set(LPOPER pr)
{
#pragma XLLEXPORT
	HANDLEX result = INVALID_HANDLEX;

	try {
		result = handles::insert(new OPER(*pr));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return result;
}

AddIn xai_range_get(
	Function(XLL_LPOPER, L"xll_range_get", L"RANGE")
	.Arguments({
		Arg(XLL_HANDLEX, L"handle", L"is a handle to a range.")
		})
	.Category(L"XLL")
	.FunctionHelp(L"Return a range given a handle.")
);
LPXLOPER12 WINAPI xll_range_get(HANDLEX h)
{
#pragma XLLEXPORT

	try {
		handle<OPER> h_(h);
		if (h_) {
			return h_.ptr();
		}
		OPER* po = handles::find(h);
		if (po) {
			return po;
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return const_cast<LPXLOPER12>(&ErrNA);
}