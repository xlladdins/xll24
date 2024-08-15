// depends.cpp - Force calculation dependencies.
#include "xll.h"

using namespace xll;

AddIn xai_depends(
	Function(XLL_LPXLOPER, "xll_depends", "DEPENDS")
	.Arguments({
		Arg(XLL_LPXLOPER, "cell", "is reference to a cell."),
		Arg(XLL_LPXLOPER, "dependence", "is reference to cell dependency."),
		})
	.FunctionHelp("Return cell after dependency is called.")
	.Category("XLL")
);
LPXLOPER12 WINAPI xll_depends(LPXLOPER12 c, LPXLOPER12)
{
#pragma XLLEXPORT

	return c;
}