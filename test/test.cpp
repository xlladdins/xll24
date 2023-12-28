// test.cpp
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
		ensure(OPER(true).as_num() == 1);
		ensure(OPER(123).as_num() == 123);
		ensure(OPER(1.23).as_num() == 1.23);
		ensure(as_num(Missing) == 0);
		ensure(as_num(Nil) == 0);
		ensure(_isnan(OPER(L"abc").as_num()));
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
		OPER o2{ o };
		ensure(o == o2);
		o = o2;
		ensure(!(o != o2));
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
		ensure((OPER(L"abc") & OPER(L"xyz")) == L"abcxyz");
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
		ensure(o.xltype == xltypeMulti);
		ensure(rows(o) == 1);
		ensure(columns(o) == 3);
		ensure(o == OPER({ OPER(1.),OPER(2.),OPER(3.) }));
	}
	{
		OPER o = Evaluate(OPER(L"{1, \"a\"; TRUE, #N/A}"));
		ensure(o.xltype == xltypeMulti);
		ensure(rows(o) == 2);
		ensure(columns(o) == 2);
		ensure(o == OPER({ OPER(1.),OPER(L"a"),OPER(true),OPER(xlerr::NA) }).reshape(2, 2));
	}
	{
		OPER o = Evaluate(OPER(L"=Sheet1!$A$1"));
		ensure(o == ErrRef);
	}
	{
		OPER o = Evaluate(OPER(L"=R[1]C[1]"));
		ensure(o == ErrValue);
	}

	return 0;
}

int err_test()
{
	{ // Errors are impervious to Text and Evaluate
#define ERR_TEST(a,b,c) ensure(Text(OPER(xlerr::##a)) == Err##a);
		XLL_TYPE_ERR(ERR_TEST);
#undef ERR_TEST
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
		ensure(!(m != m))
		ensure(rows(m) == 1);
		ensure(columns(m) == 2);
		ensure(size(m) == 2);
		ensure(m[0] == 1.23);
		ensure(m[1] == L"abc");
		m[0] = m;
		m[1] = m;
		m[1][1] = m;
		OPER m2{ m };
		ensure(m == m2);
		m = m2;
		ensure(!(m != m2));
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
	return 0;
}

int evaluate_test()
{
	{
		OPER o = Evaluate(OPER(L"COUNT(!A1:B2)"));
		o = o;
	}
	return 0;
}

int excel_test()
{
	{
		OPER o = Excel(xlSheetId);
		ensure(o.xltype == xltypeNum);
		ensure(o.val.num == 1);
	}

	return 0;
}

int xll_test()
{
	set_alert_level(7);
	//XlfRegister(Macro(L"?xll_test", L"XLL.TEST"));
	try {
		utf8::test();
		num_test();
		str_test();
		err_test();
		bool_test();
		multi_test();
		evaluate_test();
		excel_test();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;
}

Auto<Open> xao_test(xll_test);
//AddIn xai_test(Macro(L"xll_test", L"XLL.TEST"));
#endif