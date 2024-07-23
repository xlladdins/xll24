// paste.cpp - Paste shortcuts.
#include "xll.h"

using namespace xll;

// Arg pointer from function text in cell.
const Args* args(const OPER& text)
{
	const auto c = AddIn::addins.find(text);

	return c == AddIn::addins.end() ? nullptr : c->second;
}

// Paste function with default arguments.
AddIn xai_pasteb(
	Macro("xll_pasteb", "XLL.PASTEB")
);
int WINAPI xll_pasteb()
{
#pragma XLLEXPORT
	int result = TRUE;

	try {
		OPER caller = Excel(xlfActiveCell);
		OPER text = Excel(xlCoerce, caller);
		const Args* pargs = args(text);
		ensure(pargs || !"xll_pasteb: add-in not found");

		text &= OPER(L"(");
		OPER comma(L"");
		for (int i = 0; i < pargs->count(); ++i) {
			text &= comma;
			text &= Excel(xlfText, pargs->argumentInit[i], L"General");
			
			comma = L", ";
		}
		text &= OPER(L")");
		text = OPER("=") & text;

		Excel(xlcFormula, text, caller);
	}
	catch (const std::exception& ex) {
		result = FALSE;
		XLL_ERROR(ex.what());
	}
	catch (...) {
		result = FALSE;
		XLL_ERROR("xll_pasteb: unknown exception");
	}

	return result;
}
On<xlcOnKey> xok_pasteb(ON_CTRL ON_SHIFT "B", "XLL.PASTEB");

// Paste function above default arguments.
AddIn xai_pastec(
	Macro("xll_pastec", "XLL.PASTEC")
);
int WINAPI xll_pastec()
{
#pragma XLLEXPORT
	int result = TRUE;

	try {
		OPER caller = Excel(xlfActiveCell);
		OPER text = Excel(xlCoerce, caller);
		const Args* pargs = args(text);
		ensure(pargs || !"xll_pastec: add-in not found");

		text &= OPER(L"(");
		OPER comma(L"");
		for (int i = 0; i < pargs->count(); ++i) {
			text &= comma;
			OPER calleri = Excel(xlfOffset, caller, i + 1, 0);
			text &= Excel(xlfRelref, calleri, caller);
			Excel(xlSet, calleri, pargs->argumentInit[i]);

			comma = L", ";
		}
		text &= OPER(L")");
		text = OPER("=") & text;

		Excel(xlcFormula, text, caller);
	}
	catch (const std::exception& ex) {
		result = FALSE;
		XLL_ERROR(ex.what());
	}
	catch (...) {
		result = FALSE;
		XLL_ERROR("xll_pasteb: unknown exception");
	}

	return result;
}
On<xlcOnKey> xok_pastec(ON_CTRL ON_SHIFT "C", "XLL.PASTEC");

auto AlignRight(const OPER& ref)
{
	OPER ac = Excel(xlfActiveCell);

	Excel(xlcSelect, ref);
	Excel(xlcAlignment, OPER(4));
	Excel(xlcSelect, ac);

	return ac;
}
auto Bold(const OPER& ref)
{
	return ref;
}

// Paste function and define names.
AddIn xai_pasted(
	Macro("xll_pasted", "XLL.PASTED")
);
int WINAPI xll_pasted()
{
#pragma XLLEXPORT
	int result = TRUE;

	try {
		OPER caller = Excel(xlfActiveCell);
		OPER text = Excel(xlCoerce, caller);
		const Args* pargs = args(text);
		ensure(pargs || !"xll_pasted: add-in not found");

		AlignRight(caller);

		OPER formula = OPER(L"=") & text & OPER(L"(");
		OPER comma(L"");
		for (int i = 0; i < pargs->count(); ++i) {
			formula &= comma;
			OPER texti = Excel(xlfOffset, caller, i + 1, 0);
			const OPER& name = pargs->argumentName[i];
			formula &= name;
			Excel(xlSet, texti, name);
			AlignRight(texti);
			OPER calleri = Excel(xlfOffset, caller, i + 1, 1);
			Excel(xlSet, calleri, pargs->argumentInit[i]);
			Excel(xlcDefineName, name, calleri);

			comma = L", ";
		}
		formula &= OPER(L")");

		const OPER caller1 = Excel(xlfOffset, caller, 0, 1);
		Excel(xlcFormula, formula, caller1);
		Excel(xlcDefineName, text, caller1);
		Excel(xlcSelect, caller1);
	}
	catch (const std::exception& ex) {
		result = FALSE;
		XLL_ERROR(ex.what());
	}
	catch (...) {
		result = FALSE;
		XLL_ERROR("xll_pasteb: unknown exception");
	}

	return result;
}
On<xlcOnKey> xok_pasted(ON_CTRL ON_SHIFT "D", "XLL.PASTED");
