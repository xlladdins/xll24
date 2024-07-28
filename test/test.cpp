// test.cpp
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
//#if 0
#include <numeric>
#include <random>
#include "../xll.h"
#include "../excel_time.h"

using namespace xll;

int excel_time_test()
{
	{
		time_t t = time(nullptr); // UTC time
		ensure(t != -1);
		const auto t_ = to_time_t(from_time_t(t));
		ensure(fabs(t - t_) <= 1);
		
		auto now = Excel(xlfNow); // local time
		OPER d(from_time_t(t));
		ensure(Excel(xlfYear, now) == Excel(xlfYear, d));
		ensure(Excel(xlfMonth, now) == Excel(xlfMonth, d));
		ensure(Excel(xlfDay, now) == Excel(xlfDay, d));
		ensure(Excel(xlfHour, now) == Excel(xlfHour, d));
		ensure(Excel(xlfMinute, now) == Excel(xlfMinute, d));
		auto ds = Num(Excel(xlfSecond, now)) - Num(Excel(xlfSecond, d));
		ensure(fabs(ds) <= 1);
	}

	return 0;
}

//#ifdef _DEBUG

// _crtBreakAlloc = 0;
AddIn xai_id(
	Function(XLL_LPOPER, L"xll_id", L"XLL.ID")
	.Arguments({
		Arg(XLL_LPOPER, L"ref", L"is a reference to a cell."),
		})
);
LPOPER WINAPI xll_id(LPOPER pref)
{
#pragma XLLEXPORT
	return pref;
}

XLL_CONST(DOUBLE, XLL_PI, 3.14159265358979323846, "Return the constant pi.", "XLL", "");

//#if 0
int ref_test()
{
	XLREF12 r12 = { .rwFirst = 1, .rwLast = 2, .colFirst = 3, .colLast = 4 };
	{
		ensure(rows(r12) == 2);
		ensure(columns(r12) == 2);
		ensure(size(r12) == 4);
		ensure(r12 == r12);
		ensure(!(r12 != r12));
		REF r{ r12 };
		ensure(r == r12);
		auto s = r;
		ensure(!(s != r));
	}

	return 0;
}

int num_test()
{
	{
		OPER o(1.23);
		ensure(type(o) == xltypeNum);
		ensure(o.val.num == 1.23);
		ensure(o == 1.23);
		o = 3.21;
		ensure(type(o) == xltypeNum);
		ensure(o.val.num == 3.21);
		ensure(o == 3.21);
		OPER o2{ o };
		ensure(o == o2);
		o = o2;
		ensure(!(o != o2));
	}
	{
		ensure(asNum(OPER(true)) == 1);
		ensure(asNum(OPER(123)) == 123);
		ensure(asNum(OPER(1.23)) == 1.23);
		ensure(asNum(Missing) == 0);
		ensure(asNum(Nil) == 0);
		ensure(_isnan(asNum(OPER(L"abc"))));
	}

	return 0;
}

int str_test()
{
	{
		constexpr XLOPER s = Str("\03abc");
		static_assert(s.xltype == xltypeStr);
		static_assert(s.val.str[0] == 3);
	}
	{
		constexpr XLOPER12 s = Str(L"\03abc");
		static_assert(s.xltype == xltypeStr);
		static_assert(s.val.str[0] == 3);
	}
	OPER General(L"General");
	{
		OPER o(L"");
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str && o.val.str[0] == 0);
		ensure(o == L"");
		ensure(o == "");
		OPER o2{ o };
		ensure(o == o2);
		o = o2;
		ensure(!(o != o2));

		o = L"abc";
		ensure(o.xltype == xltypeStr);
		ensure(o == L"abc");

		o = "def"; // converts to wide character string
		ensure(o.xltype == xltypeStr);
		ensure(o == L"def");
	}
	{
		OPER o("");
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str && o.val.str[0] == 0);
		ensure(o == "");
		ensure(o == L"");
	}
	{
		ensure(OPER(L"").xltype == xltypeStr);
		ensure(OPER(L"").val.str[0] == 0);
		ensure(OPER(L"abc").xltype == xltypeStr);
		ensure(OPER(L"abc").val.str[0] == 3);
		ensure(OPER(L"abc").val.str[3] == L'c');

		ensure((OPER(L"a") != OPER(L"b")));
		ensure((OPER(L"a") == OPER(L"a")));
		ensure((OPER(L"b") != OPER(L"a")));
		ensure((OPER(L"aa") != OPER(L"a")));
	}
	{
		OPER o("abc");
		ensure(o == L"abc");
		o = o & OPER(L"def");
		ensure(o == L"abcdef");
		ensure((OPER(L"abc") & OPER(L"xyz")) == OPER(L"abcxyz"));
	}
	{
		OPER o = Excel(xlfText, 1.23, General);
		ensure(o == L"1.23");
		ensure(Excel(xlfValue, o) == 1.23);
	}
	{
		OPER o = Excel(xlfText, L"abc", General);
		ensure(o == L"abc");
	}
	{
		OPER o = Excel(xlfText, true, General);
		ensure(o == L"TRUE");
		ensure(Excel(xlfValue, o) == ErrValue);
	}
	{
		OPER o = Excel(xlfText, 123, General);
		ensure(o == L"123");
		ensure(Excel(xlfValue, o) == 123.0); // Int gets converted to Num
	}
	{
		OPER o = Excel(xlfEvaluate, L"{1,2,3}");
		ensure(type(o) == xltypeMulti);
		//ensure(o.xltype & xlbitXLFree);
		ensure(rows(o) == 1);
		ensure(columns(o) == 3);
		ensure(o == OPER({ OPER(1.),OPER(2.),OPER(3.) }));
	}
	{
		OPER o = Excel(xlfEvaluate, L"{1, \"a\"; TRUE, #N/A}");
		ensure(type(o) == xltypeMulti);
		//ensure(o.xltype & xlbitXLFree);
		ensure(rows(o) == 2);
		ensure(columns(o) == 2);
		ensure(o == OPER({ OPER(1.),OPER(L"a"),OPER(true),OPER(xlerr::NA) }).reshape(2, 2));
	}
	{
		//OPER o = Evaluate(OPER(L"=Sheet1!$A$1"));
		//ensure(o == ErrRef);
	}
	{
		//OPER o = Evaluate(OPER(L"=R[1]C[1]"));
		//ensure(o == ErrValue);
	}

	return 0;
}

int err_test()
{
	{
		// Errors are impervious to Text and Evaluate
#define ERR_TEST(a,b,c) ensure(Excel(xlfText, xlerr::##a, OPER(L"General")) == Err##a);
		XLL_TYPE_ERR(ERR_TEST);
#undef ERR_TEST

		// Evaluate converts string error to XLOPER.
#define ERR_TEST(a,b,c) ensure(Excel(xlfEvaluate, b) == Err##a);
		XLL_TYPE_ERR(ERR_TEST)
#undef ERR_TEST
	}

	return 0;
}

int bool_test()
{
	{
		OPER b(true);
		ensure(b.xltype == xltypeBool);
		ensure(b.val.xbool == TRUE);
		ensure(b == true);
		b = false;
		ensure(b.xltype == xltypeBool);
		ensure(b.val.xbool == FALSE);
		ensure(b == false);
	}
	{
		OPER b(L"abc");
		ensure(b != true);
		ensure(b != false);
		b = true;
		ensure(b == true);
	}

	return 0;
}

int multi_test()
{
	{
		OPER m({ OPER(1.23), OPER(L"abc") });
		ensure(type(m) == xltypeMulti);
		ensure(m == m);
		ensure(!(m != m));
		ensure(rows(m) == 1);
		ensure(columns(m) == 2);
		ensure(size(m) == 2);
		ensure(m[0] == 1.23);
		ensure(m[1] == L"abc");

		OPER m0 = m;
		m[0] = m;
		ensure(m[0] == m0);

		m[1] = m;
		ensure(m[0] == m0);
		ensure(m[1][0] == m0);
		ensure(m[1][1] == L"abc");
		
		OPER m2{ m };
		ensure(m == m2);
		m = m2;
		ensure(!(m != m2));

		m[0][0] = m;
		m[0][0][0] = m;
	}
	{
		OPER m;
		m.reshape(2, 3);
		m[0] = 0;
		ensure(m[0] == 0);
		m(0, 0) = 0;
		ensure(m[0] == 0);
		m(0, 0) = 0;
		ensure(m[0] == 0);

		m(0, 1) = 2;
		ensure(m[1] == 2);
	}
	{
		OPER o(L"abc");
		ensure(type(o) == xltypeStr);
		ensure(view(o) == L"abc");
		o = L"de";
		ensure(type(o) == xltypeStr);
		ensure(o == L"de");
		o.enlist();
		ensure(type(o) == xltypeMulti);
		ensure(rows(o) == 1);
		ensure(columns(o) == 1);
		ensure(o[0] == L"de");
	}
	{
		OPER o({ OPER(L"a"),OPER(L"b") });
		o[0] = o;
		o[0][0] = o;
		o[0][0][0] = o;
	}
	{
		OPER o(L"abc");
		o.enlist();
		ensure(o[0] == L"abc");
		o.enlist();
		ensure(o[0][0] == L"abc");
		o.enlist();
		ensure(o[0][0][0] == L"abc");
	}
	{
		static std::mt19937 gen;
		const auto rand_between = [](int a, int b) { // uniform int in [a, b]
			std::uniform_int_distribution<int> dis(a, b);

			return dis(gen);
		};
		const auto rand_multi = [&](int a, int b) {
			return OPER(rand_between(a, b), rand_between(a, b), nullptr);
		};
		const auto rand_assign = [&](OPER& o) {
			auto r = rows(o);
			auto c = columns(o);
			OPER o_ = rand_multi(r, c);
			auto i = rand_between(0, r - 1);
			auto j = rand_between(0, c - 1);
			o(i, j) = o_;
		};
		OPER o = rand_multi(100, 200);
		for (int i = 0; i < 100; ++i) {
			rand_assign(o);
		}
	}
	{
		OPER o({ OPER(1.23), OPER(L"abc"), OPER(true) });
		ensure(o == OPER().push_bottom(o));
		ensure(o == o.push_bottom(OPER()));
		ensure(OPER(L"abc") == OPER(L"abc").push_bottom(OPER()));
		ensure(OPER(L"abc") == OPER().push_bottom(OPER(L"abc")));
	}
	{
		OPER o({ OPER(1.23), OPER(L"abc"), OPER(true) });
		OPER o2 = o.push_bottom(o);
		ensure(rows(o2) == 2);
		ensure(columns(o2) == 3);
		ensure(o2[0] == o[0]);
		ensure(o2[1] == o[1]);
		ensure(o2[2] == o[2]);
		ensure(o2[3] == o[0]);
		ensure(o2[4] == o[1]);
		ensure(o2[5] == o[2]);
	}
	{
		OPER o({ OPER(1.23), OPER(L"abc"), OPER(true) });
		o[2] = o;
		OPER o2 = o.push_bottom(o);
		ensure(rows(o2) == 2);
		ensure(columns(o2) == 3);
		ensure(o2[0] == o[0]);
		ensure(o2[1] == o[1]);
		ensure(o2[2] == o[2]);
		ensure(o2[3] == o[0]);
		ensure(o2[4] == o[1]);
		ensure(o2[5] == o[2]);

		ensure(o2[1] == index(o2, 1));
		ensure(o2(1, 1) == index(o2, 1, 1));
	}
	{
		OPER o({ OPER(1),OPER(2) });
		o.append(OPER(3));
		ensure(o == OPER({ OPER(1),OPER(2), OPER(3)}));
	}
	{
		OPER o({ OPER(1),OPER(2) });
		o.append(OPER());
		ensure(o == OPER({ OPER(1),OPER(2), OPER() }));
	}

	return 0;
}

int markup_test()
{
	{
		ensure(AddIn::addins.contains(OPER(L"XLL.HYPOT")));
		auto args = AddIn::addins[OPER(L"XLL.HYPOT")];
		auto mu = args->markup();
		ensure(rows(mu) == 14);
		ensure(columns(mu) == 2);
		ensure(mu(10, 0) == L"argumentHelp");
		
	}
	return 0;
}

int evaluate_test()
{
	{
		OPER o = Excel(xlfEvaluate, L"=1+2");
		ensure(o != L"=1+2");
		ensure(o == 3.);
		ensure(o == 3);
		ensure(o != true);
	}
	{
		OPER o = Excel(xlfEvaluate, "=1+2"); // utf-8 to wide character string
		ensure(o != "=1+2");
		ensure(o == 3.);
		ensure(o == 3);
		ensure(o != true);
	}
	{
		OPER o = Excel(xlfEvaluate, L"SUM(1,2)");
		ensure(o == 3);
	}
	{
		OPER o = Excel(xlfEvaluate, "SUM(1,2)");
		ensure(o == 3);
	}
	{
		OPER o = Excel(xlfEvaluate, OPER(1.23));
		ensure(o == 1.23);
	}
	{
		OPER o = Excel(xlfEvaluate, OPER(L"abc"));
		ensure(o == ErrName);
	}
	{
		OPER o = Excel(xlfEvaluate, OPER(true));
		ensure(o == true);
	}


	return 0;
}

int excel_test()
{
	{
		OPER o = Excel(xlfToday);
		OPER p = Excel(xlfDate, Excel(xlfYear, o), Excel(xlfMonth, o), Excel(xlfDay, o));
		ensure(o == p);
	}
	{
		OPER o = Excel(xlfDate, 2024, 1, 2);
		OPER p = Excel(xlfText, o, L"yyyy-mm-dd");
		ensure(p == L"2024-01-02");
	}

	return 0;
}

int fp_test()
{
	{
		FPX a(2, 3);
		ensure(size(a) == 6);
		ensure(rows(a) == 2);
		ensure(columns(a) == 3);
		ensure(a == a);
		ensure(!(a != a));
		a[0] = 0;
		ensure(a[0] == 0);
		a(0, 1) = 0;
		ensure(a[1] == 0);
	}
	{
		FPX a({ 1,2,3 });
		ensure(size(a) == 3);
		ensure(rows(a) == 1);
		ensure(columns(a) == 3);
		ensure(a[0] == 1);
		ensure(a[1] == 2);
		ensure(a[2] == 3);
		a.resize(2, 2);
		ensure(size(a) == 4);
		ensure(rows(a) == 2);
		ensure(columns(a) == 2);
		ensure(a[0] == 1);
		ensure(a[1] == 2);
		ensure(a[2] == 3);
		ensure(a(1, 0) == 3);

		FPX a2{ a };
		ensure(a == a2);
		a = a2;
		ensure(!(a != a2));
	}
	{
		FPX a;
		ensure(size(a) == 0);
		ensure(rows(a) == 0);
		ensure(columns(a) == 0);
		a.push_back(1.1);
		ensure(size(a) == 1);
		ensure(rows(a) == 1);
		ensure(columns(a) == 1);
		ensure(a[0] == 1.1);

		a.resize(0, 0);
		ensure(size(a) == 0);
		ensure(rows(a) == 0);
		ensure(columns(a) == 0);

		a.resize(2, 3);
		ensure(size(a) == 6);
		ensure(rows(a) == 2);
		ensure(columns(a) == 3);
		double x[] = { 7,8,9 };
		a.vstack(FPX(x));
		ensure(rows(a) == 3);
		ensure(columns(a) == 3);
		ensure(a[6] == 7);
		ensure(a[7] == 8);
		ensure(a[8] == 9);
	}
	{
		FPX a({ 1,2,3 });
		FPX b({ 4,5,6 });
		auto c = a.hstack(b);
		ensure(rows(c) == 1);
		ensure(columns(c) == 6);
		ensure(c[0] == 1);
		ensure(c[1] == 2);
		ensure(c[2] == 3);
		ensure(c[3] == 4);
		ensure(c[4] == 5);
		ensure(c[5] == 6);

		c.hstack(c);
		ensure(rows(c) == 1);
		ensure(columns(c) == 12);
		ensure(c[0] == 1);
		ensure(c[1] == 2);
		ensure(c[2] == 3);
		ensure(c[3] == 4);
		ensure(c[4] == 5);
		ensure(c[5] == 6);
		ensure(c[6] == 1);
		ensure(c[7] == 2);
		ensure(c[8] == 3);
		ensure(c[9] == 4);
		ensure(c[10] == 5);
		ensure(c[11] == 6);
	}

	return 0;
}

int int_test()
{
	{
		OPER o;
		o = 1;
		ensure(o == 1);
		ensure(type(o) == xltypeInt);
		o = 0;
		ensure(o == 0);
		ensure(type(o) == xltypeInt);
	}
	return 0;
}

int WINAPI xll_test()
{
#pragma XLLEXPORT
	int xal = get_alert_mask();
	set_alert_mask(1);
	int al = get_alert_mask();
	ensure(1 == al);
	set_alert_mask(xal);

	AddInManagerInfo(OPER("The xll_test add-in"));
	Macro m(L"?xll_test", L"XLL.TEST");
	XlfRegister(&m);

	try {
		utf8::test();
		ref_test();
		num_test();
		str_test();
		int_test();
		err_test();
		bool_test();
		multi_test();
		markup_test();
		evaluate_test();
		excel_test();
		fp_test();
		excel_time_test();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return false;
	}

	return true;
}

Auto<OpenAfter> xao_test(xll_test);
//#endif

Auto<Open> xao_sam([]() { set_alert_mask(7); return true; });

const AddIn xai_const(Function(XLL_DOUBLE, "xll_const", "XLL.CONST"));
double WINAPI xll_const()
{
#pragma XLLEXPORT
	return 3.14159265358979323846;
}

//AddIn xai_test(Macro(L"xll_test", L"XLL.TEST"));
///*
const AddIn xai_hypot(Function(XLL_DOUBLE, L"xll_hypot", L"XLL.HYPOT")
	.Arguments({
		Arg(XLL_DOUBLE, L"x", L"is a number."),
		Arg(XLL_DOUBLE, L"y", L"is a number.", "=2 + 2"),
		})
	.Category(L"XLL")
	.FunctionHelp("Return the length of the hypotenuse of a right triangle with sides x and y.")
	.HelpTopic("https://learn.microsoft.com/en-us/cpp/c-runtime-library/reference/hypot-hypotf-hypotl-hypot-hypotf-hypotl?view=msvc-170")
//	.Documentation(L"Optional documentation.")
);
double WINAPI xll_hypot(double x, double y)
{
#pragma XLLEXPORT
	const OPER o = Excel(xlfSqrt, Excel(xlfSumsq, OPER(x), OPER(y)));
	double h = std::hypot(x, y);
	if (h != Num(o)) {
		XLL_WARNING("This should never happen.");
	}
	if (o != h) {
		XLL_WARNING("This should never happen.");
	}

	return h;
}
//*/
AddIn xai_array(
	Function(XLL_FP, L"xll_array", L"XLL.ARRAY")
	.Arguments({
		Arg(XLL_FP, L"array", L"is an array of numbers.", "={1,2,3;4,5,6}"), // row default value
		Arg(XLL_DOUBLE, L"scalar", L"is a scalar.", 2.71828), // scalar default value
		})
	.Category(L"XLL")
	.FunctionHelp("Return the sum of all the numbers passed in.")
	.HelpTopic("https://docs.microsoft.com/en-us/cpp/standard-library/accumulate?view=msvc-170")
);
_FP12* WINAPI xll_array(_FP12* pa, double s)
{
#pragma XLLEXPORT
	pa->array[0] = s;
	return pa;
}

AddIn xai_relref(
	Function(XLL_LPOPER, L"xll_relref", L"XLL.RELREF")
	.Arguments({
		Arg(XLL_LPXLOPER, L"ref", L"is the cell or cells to which you want to create a relative reference.."),
		Arg(XLL_LPXLOPER, L"rel", L"is the cell from which you want to create the relative reference."),
		})
	.Uncalced()
	.Category(L"XLL")
	.FunctionHelp(L"Returns the reference of a cell or cells relative to the upper-left cell of rel_to_ref. "
		L"The reference is given as an R1C1-style relative reference in the form of text, such as \"R[1]C[1]\".")
);
LPOPER WINAPI xll_relref(LPXLOPER12 pref, LPXLOPER12 prel)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		ensure(isSRef(*pref) || isRef(*pref));
		ensure(isSRef(*prel) || isRef(*prel));

		o = Excel(xlfRelref, *pref, *prel);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_accumulate(
	Function(XLL_DOUBLE, L"xll_accumulate", L"XLL.ACCUMULATE")
	.Arguments({
		Arg(XLL_FP, L"array", L"is an array of numbers."),
		})
	.Category(L"XLL")
	.FunctionHelp("Return the sum of all the numbers passed in.")
	.HelpTopic("https://docs.microsoft.com/en-us/cpp/standard-library/accumulate?view=msvc-170")
);
double WINAPI xll_accumulate(_FP12* pa)
{
#pragma XLLEXPORT
	return std::accumulate(span(*pa).begin(), span(*pa).end(), 0.);
}

AddIn xai_get_workspace(
	Function(XLL_LPOPER, "xll_get_workspace", "XLL.GET.WORKSPACE")
	.Arguments({
			Arg(XLL_LPOPER, "type_num", " is a number that specifies what type of workspace information you want."),
		})
	.Uncalced()
	.Category(L"XLL")
	.FunctionHelp(L"Returns information about a workbook.")
	.HelpTopic(L"https://xlladdins.github.io/Excel4Macros/get.workspace.html")
);
LPOPER WINAPI xll_get_workspace(LPOPER po)
{
#pragma XLLEXPORT
	static OPER o;
	
	o = Excel(xlfGetWorkspace, *po);

	return &o;
}

AddIn xai_get_workbook(
	Function(XLL_LPOPER, L"xll_get_workbook", L"XLL.GET.WORKBOOK")
	.Arguments({
		Arg(XLL_LPOPER, L"type_num", L" is a number that specifies what type of workbook information you want."),
	})
	.Uncalced()
	.Category(L"XLL")
	.FunctionHelp(L"Returns information about a workbook.")
	.HelpTopic(L"https://xlladdins.github.io/Excel4Macros/get.workbook.html")
);
LPOPER WINAPI xll_get_workbook(LPOPER po)
{
#pragma XLLEXPORT
	static OPER o;

	o = Excel(xlfGetWorkbook, *po);

	return &o;
}
/*
AddIn xai_evaluate(
	Function(XLL_LPOPER, L"xll_evaluate", L"XLL.EVALUATE")
	.Arguments({
		Arg(XLL_LPOPER, L"formula_text", L"is the expression in the form of text that you want to evaluate."),
		})
	.Category(L"XLL")
	.FunctionHelp(L"Evaluates a formula or expression that is in the form of text and returns the result.")
	.HelpTopic(L"https://xlladdins.github.io/Excel4Macros/evaluate.html")
);
LPOPER WINAPI xll_evaluate(LPOPER po)
{
#pragma XLLEXPORT
	static OPER o;

	o = Excel(xlfEvaluate, *po);

	return &o;
}
*/
//#endif // 0
//#endif // 0

XLL_CONST(LPOPER, XLL_NA, (LPOPER)& ErrNA, "Return the #N/A error value.", "XLL", "");

double my_double(double x)
{
	return 2*x;
}

XLL_CONST(HANDLEX, MY_DOUBLE, safe_handle(my_double), "Return the handle data type.", "XLL", "");

AddIn xai_my_double(
	Function(XLL_DOUBLE, L"xll_my_double", L"XLL.MY.DOUBLE")
	.Arguments({
		Arg(XLL_HANDLEX, L"handle", L"is a handle to my_double."),
		Arg(XLL_DOUBLE, L"x", L"is a number."),
		})
	.Category(L"XLL")
	.FunctionHelp(L"Return 2 times the number.")
);
double WINAPI xll_my_double(HANDLEX h, double x)
{
#pragma XLLEXPORT
	auto h_ = safe_value<double(*)(double)>(h);

	return h_(x);
}
