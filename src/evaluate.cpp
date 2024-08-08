// evaluate.cpp - Call xlfEvaluate.
#include "xll.h"

using namespace xll;

AddIn xai_evaluate(
	Function(XLL_LPOPER, "xll_evaluate", "EVAL")
	.Arguments({
		Arg(XLL_LPOPER, "formula", "is a formula to evaluate.")
		})
	.FunctionHelp("Return the value of the formula.")
	.Category("XLL")
);
LPXLOPER12 WINAPI xll_evaluate(LPXLOPER12 p)
{
#pragma XLLEXPORT
	try {
		ensure(0 == ::Excel12(xlfEvaluate, p, 1, p));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return XLFree(*p);
}