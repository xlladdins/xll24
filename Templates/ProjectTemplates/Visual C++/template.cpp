// template.cpp
#include <cmath>
#include "template.h"

using namespace xll;

AddIn xai_hermite(
	Function(XLL_DOUBLE, "xll_hermite", "XLL.HERMITE")
	.Arguments({
		Arg(XLL_WORD, "n", "the degree of the polynomial"),
		Arg(XLL_DOUBLE, "x", "the argument, a floating-point value."),
	})
	.Category(CATEGORY)
	.FunctionHelp("Return the value of the order-n Hermite polynomial at x.")
	.HelpTopic("https://en.cppreference.com/w/cpp/numeric/special_functions/hermite")
);
double WINAPI xll_hermite(UINT n, double x)
{
#pragma XLLEXPORT
	return std::hermite(n, x);	
}