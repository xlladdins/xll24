// test.cpp
#include "../xll.h"

using namespace xll;

//Auto<Open> xao_test(xll::test);

int xll_test()
{
	try {
		{
			OPER o;
			ensure(o.xltype == xltypeNil);
			OPER o2{ o };
			ensure(o == o2);
			o = o2;
			ensure(!(o != o2));
		}
		{
			OPER o;
			o = 1.23;
			ensure(type(o) == xltypeNum);
			ensure(o.val.num == 1.23);
			ensure(o == 1.23);
		}
		{
			/*
			static_assert(Str().xltype == xltypeStr); // zero length string
			static_assert(Str().val.str[0] == 0);
			static_assert(Str(L"").xltype == xltypeStr);
			static_assert(Str(L"").val.str[0] == 0);
			static_assert(Str(L"\03" L"abc").xltype == xltypeStr);
			static_assert(Str(L"\03" L"abc").val.str[0] == 3);
			static_assert(Str(L"\03" L"abc").val.str[3] == L'c');

			static_assert(Str() == Str());
			static_assert(Str(L"") == Str(L""));
			static_assert((Str(L"\001a") != Str(L"\001b")));
			static_assert((Str(L"\001a") == Str(L"\001a")));
			static_assert((Str(L"\001b") != Str(L"\001a")));
			static_assert((Str(L"\002aa") != Str(L"\001a")));
			*/
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