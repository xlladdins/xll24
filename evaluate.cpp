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
	static OPER result;

	try {
		result = Excel(xlfEvaluate, *p);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &result;
}