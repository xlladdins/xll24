// template.cpp
#include "xll.h"

using namespace xll;

AddIn xai_function(
	Function(XLL_DOUBLE, L"xll_function", L"XLL.FUNCTION")
	.Arguments({
		Arg(XLL_DOUBLE, L"x", L"is a number."),
	})
	.Category(L"XLL")
	.FunctionHelp(L"Return x.")
);
double WINAPI xll_function(double x)
{
#pragma XLLEXPORT
	return x;	
}