// test.cpp
#include <numeric>
#include "../xll.h"

using namespace xll;

#ifdef _DEBUG
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
		ensure(as_num(OPER(true)) == 1);
		ensure(as_num(OPER(123)) == 123);
		ensure(as_num(OPER(1.23)) == 1.23);
		ensure(as_num(Missing) == 0);
		ensure(as_num(Nil) == 0);
		ensure(_isnan(as_num(OPER(L"abc"))));
	}

	return 0;
}

int str_test()
{
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
		o &= OPER(L"def");
		ensure(o == L"abcdef");
		ensure((OPER(L"abc") & OPER(L"xyz")) == OPER(L"abcxyz"));
	}
	{
		OPER o = Text(Nil);
		ensure (o == L"0")
	}
	{
		OPER o = Text(Missing);
		ensure(o == L"0");
	}
	{
		OPER o = Text(OPER(1.23));
		ensure(o == L"1.23");
		ensure(Value(o) == 1.23);
	}
	{
		OPER o = Text(OPER(L"abc"));
		ensure(o == L"abc");
	}
	{
		OPER o = Text(OPER(true));
		ensure(o == L"TRUE");
		ensure(Value(o) == ErrValue);
	}
	{
		OPER o = Text(OPER(123));
		ensure(o == L"123");
		ensure(Value(o) == 123.0); // Int gets converted to Num
	}
	{
		OPER o = Evaluate(OPER(L"{1,2,3}"));
		ensure(type(o) == xltypeMulti);
		//ensure(o.xltype & xlbitXLFree);
		ensure(rows(o) == 1);
		ensure(columns(o) == 3);
		ensure(o == OPER({ OPER(1.),OPER(2.),OPER(3.) }));
	}
	{
		OPER o = Evaluate(OPER(L"{1, \"a\"; TRUE, #N/A}"));
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
#define ERR_TEST(a,b,c) ensure(Text(OPER(xlerr::##a)) == Err##a);
		XLL_TYPE_ERR(ERR_TEST);
#undef ERR_TEST

		// Evaluate converts string error to XLOPER.
#define ERR_TEST(a,b,c) ensure(Evaluate(OPER(b)) == Err##a);
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
		ensure(m[1] == L"abc");

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
		OPER o({
			{ L"a", OPER(1) } ,
			{ L"b", OPER(2) } ,
			{ L"c", OPER(3) } ,
			});
		ensure(size(o) == 6);
	}
	{
		OPER o({ OPER(L"a"),OPER(L"b") });
		o[0] = o;
		o[0][0] = o;
		o[0][0][0] = o;
	}

	return 0;
}

int evaluate_test()
{
	{
		OPER o = Evaluate(OPER(L"=1+2"));
		ensure(o != L"=1+2");
		ensure(o == 3.)
		ensure(o == 3);
		ensure(o != true);
	}
	{
		OPER o = Evaluate(OPER("=1+2")); // utf-8 to wide character string
		ensure(o != "=1+2");
		ensure(o == 3.);
		ensure(o == 3);
		ensure(o != true);
	}
	{
		OPER o = Evaluate(OPER(L"SUM(1,2)"));
		ensure(o == 3);
	}
	{
		OPER o = Evaluate(OPER("SUM(1,2)"));
		ensure(o == 3);
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
		OPER p = Text(o, OPER(L"yyyy-mm-dd"));
		ensure(p == L"2024-01-02");
	}

	return 0;
}

int fp_test()
{
	{
		FPX a(2, 3);
		ensure(a.size() == 6);
		ensure(a.rows() == 2);
		ensure(a.columns() == 3);
		ensure(a == a);
		ensure(!(a != a));
	}
	{
		FPX a({ 1,2,3 });
		ensure(a.size() == 3);
		ensure(a.rows() == 1);
		ensure(a.columns() == 3);
		ensure(a[0] == 1);
		ensure(a[1] == 2);
		ensure(a[2] == 3);
		a.resize(2, 2);
		ensure(a.size() == 4);
		ensure(a.rows() == 2);
		ensure(a.columns() == 2);
		ensure(a[0] == 1);
		ensure(a[1] == 2);
		ensure(a[2] == 3);
		ensure(a(1, 0) == 3);

		FPX a2{ a };
		ensure(a == a2);
		a = a2;
		ensure(!(a != a2));
		a2 = std::move(a);
#pragma warning(push)
#pragma warning(disable: 26800)
		ensure(!a); // use after move 
		ensure(a2);

		FPX a3(std::move(a2));
		ensure(!a2); // use after move 
#pragma warning(pop)
		ensure(a3);
		ensure(a3);
	}

	return 0;
}

int xll_test()
{
	AddInManagerInfo(OPER("The xll_test add-in"));
	//XlfRegister(Macro(L"?xll_test", L"XLL.TEST"));
	///*
	Excel(xlfRegister
		, Excel(xlGetName)
		, L"xll_test"
		, Nil
		, L"XLL.TEST"
		, Nil
		, 2.0000000000000000
		, L"XLL"
		, Nil
		, Nil
		, Nil
		, L"");
	//*/
	/*
	+[0]	0x000000b5df6e7788 xltypeStr L"<C:\\Users\\keith\\source\\repos\\xlladdins\\xll\\x64\\Debug\\test.xll"	const xloper12 *
		+[1]	0x000000b5df6e76e8 xltypeStr L"\x14?xll_spreadsheet_doc"	const xloper12 *
		+[2]	0x00007ff88bc7e728 {test.xll!xll::XOPER<xloper12> o} xltypeNil	const xloper12 *
		+[3]	0x000001c327215350 xltypeStr L"\x3DOC"	const xloper12 *
		+[4]	0x00007ff88bc7e728 {test.xll!xll::XOPER<xloper12> o} xltypeNil	const xloper12 *
		+[5]	0x000001c3272152b0 xltypeNum 2.0000000000000000	const xloper12 *
		+[6]	0x000001c327214b30 xltypeStr L"\x3XLL"	const xloper12 *
		+[7]	0x00007ff88bc7e728 {test.xll!xll::XOPER<xloper12> o} xltypeNil	const xloper12 *
		+[8]	0x000000b5df6e77c8 xltypeNil	const xloper12 *
		+[9]	0x00007ff88bc7e728 {test.xll!xll::XOPER<xloper12> o} xltypeNil	const xloper12 *
		+[10]	0x000000b5df6e8998 xltypeStr L"\0"	const xloper12 *
	*/
	/*
	const auto& fargs = Function(XLL_DOUBLE, L"?xll_func", L"XLL.FUNC")
		.Arguments({
			Arg(XLL_DOUBLE, L"x", L"is a number.")
			})
		;
	XlfRegister(fargs);
	*/

	try {
		utf8::test();
		num_test();
		str_test();
		err_test();
		bool_test();
		multi_test();
		evaluate_test();
		excel_test();
		fp_test();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;
}

Auto<OpenAfter> xao_test(xll_test);
#endif

Auto<Open> xao_sal([]() { set_alert_level(7); return TRUE; });

//AddIn xai_test(Macro(L"xll_test", L"XLL.TEST"));
///*
AddIn xai_hypot(Function(XLL_DOUBLE, L"xll_hypot", L"XLL.HYPOT")
	.Arguments({
		Arg(XLL_DOUBLE, L"x", L"is a number."),
		Arg(XLL_DOUBLE, L"y", L"is a number."),
		})
	.Category(L"XLL")
	.FunctionHelp("Return the length of the hypotenuse of a right triangle with sides x and y.")
	.HelpTopic("https://learn.microsoft.com/en-us/cpp/c-runtime-library/reference/hypot-hypotf-hypotl-hypot-hypotf-hypotl?view=msvc-170")
//	.Documentation(L"Optional documentation.")
);
double WINAPI xll_hypot(double x, double y)
{
#pragma XLLEXPORT
	return std::hypot(x, y);
}
//*/

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
	Function(XLL_LPOPER, L"xll_get_workspace", L"XLL.GET.WORKSPACE")
	.Arguments({
			Arg(XLL_LPOPER, L"type_num", L" is a number that specifies what type of workspace information you want."),
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
