// test.cpp
#include "../xll.h"

using namespace xll;

int xll_test()
{
	Register(Macro(L"xll_test", L"TEST"));
	try {
		utf8::test();
		{
			OPER o;
			ensure(o.xltype == xltypeNil);
			OPER o2{ o };
			ensure(o == o2);
			o = o2;
			ensure(!(o != o2));
		}
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
			OPER m({ OPER(1.23), OPER(L"abc") });
			ensure(type(m) == xltypeMulti);
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
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;
}

Auto<Open> xao_test(xll_test);