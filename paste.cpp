// paste.cpp - Paste shortcuts.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#include "xll.h"
#include "macrofun.h"

using namespace xll;

int red;
Auto<Open> xao_ret([] {
	red = EditColor(XLL_RGB_COLOR_RED);
	return true;
});
/*
#define XLL_RGB_COLOR_ENUM(a, b, c) XLL_CONST(LONG, ENUM_COLOR_ ## a, b, "Color", "XLL Enum", "https://en.wikipedia.org/wiki/RGB_color_model");
XLL_RGB_COLOR(XLL_RGB_COLOR_ENUM)
#undef XLL_RGB_COLOR_ENUM
*/

bool isFormula(const OPER& val)
{
	return isStr(val) && view(val).starts_with(L'=');
}

bool isHandle(const OPER& val)
{
	return isStr(val) && view(val).starts_with(L'\\');
}

// Strip off leading '=' if it is a formula.
OPER Uneval(const OPER& val)
{
	return isFormula(val) ? OPER(view(val).substr(1)) : val;
}

// Construct formula from Args default values.
OPER Formula(const Args* pargs)
{
	OPER formula = OPER(L"=") & pargs->functionText & OPER(L"(");
	OPER comma(L"");
	for (int i = 0; i < pargs->count(); ++i) {
		formula &= comma;
		formula &= Excel(xlfText, Uneval(pargs->argumentInit[i]), L"General");

		comma = L", ";
	}
	formula &= OPER(L")");

	return formula;
}

// Translate by r rows and c columns.
OPER Move(const OPER& o, int r, int c)
{
	OPER sel = Excel(xlfOffset, o, r, c);
	OPER s = Excel(xlcSelect, sel);
	return sel;
}

// Expand active to size of ref.
OPER Reshape(const OPER& active, const OPER& ref)
{
	return OPER(reshape(SRef(active), rows(ref), columns(ref)));
}

// Set value and return selection.
OPER Set(const OPER& ref, const OPER& val)
{
	OPER res = ref;
	if (isFormula(val)) {
		Excel(xlcFormula, val, ref);
		const auto e = Excel(xlfEvaluate, val);
		res = Reshape(ref, e);
	}
	else {
		Excel(xlSet, ref, val);
	}

	return res;
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
		OPER active = Excel(xlfActiveCell);
		OPER text = Excel(xlCoerce, active);
		const Args* pargs = AddIn::find(text);
		ensure(pargs || !"xll_pasteb: add-in not found");

		Excel(xlcFormula, Formula(pargs), active);
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

// Paste function with default arguments below.
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
		const Args* pargs = AddIn::find(text);
		ensure (pargs || !"xll_pastec: add-in not found");
		text = pargs->functionText;

		// Expand caller to size of formula output.
		OPER output = Reshape(caller, Excel(xlfEvaluate, Formula(pargs)));
		OPER active = Move(caller, rows(output), 0);

		OPER formula = OPER(L"=") & text & OPER(L"(");
		OPER comma(L"");
		for (int i = 0; i < pargs->count(); ++i) {
			formula &= comma;

			if (isNil(pargs->argumentInit[i])) {
				formula &= Excel(xlfRelref, active, caller);
				active = Move(active, 1, 0);
			}
			else {
				const OPER ref = Set(active, pargs->argumentInit[i]);
				formula &= Excel(xlfRelref, Reshape(active, ref), caller);
				active = Move(active, rows(ref), 0);
			}

			comma = L", ";
		}
		formula &= OPER(L")");

		Excel(xlcFormula, formula, output);
		Excel(xlcSelect, caller);
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
		const Args* pargs = AddIn::find(text);
		ensure(pargs || !"xll_pasted: add-in not found");
		text = pargs->functionText;

		Excel(xlSet, caller, text);
		AlignHorizontal(Alignment::Horizontal::Right);
		FormatFontItalic();

		// Expand caller to size of formula output.
		OPER output = Reshape(Move(caller, 0, 1), Excel(xlfEvaluate, Formula(pargs)));
		OPER active = Move(caller, rows(output), 0);

		OPER formula = OPER(L"=") & text & OPER(L"(");
		OPER comma(L"");
		for (int i = 0; i < pargs->count(); ++i) {
			formula &= comma;
			const OPER& name = pargs->argumentName[i];
			formula &= name;
			Set(active, name);
			AlignHorizontal(Alignment::Horizontal::Right);
			FormatFontBold();

			active = Move(active, 0, 1);
			if (isNil(pargs->argumentInit[i])) {
				Excel(xlcDefineName, name, active);
				Excel(xlcApplyStyle, L"Input");
				active = Move(active, 1, -1);
			}
			else {
				const auto ref = Set(active, pargs->argumentInit[i]);
				Excel(xlcDefineName, name, ref);
				Excel(xlcApplyStyle, L"Input");
				active = Move(active, rows(ref), -1);
			}

			comma = L", ";
		}
		formula &= OPER(L")");


		Excel(xlcFormula, formula, output);
		Excel(xlcApplyStyle, L"Output");
		if (isHandle(text)) {
			FormatFontSize(8);
			FormatFontColor(red);
			Excel(xlcFormatNumber, "\"0x\"@");
		}
		Excel(xlcSelect, caller);
	}
	catch (const std::exception& ex) {
		result = FALSE;
		XLL_ERROR(ex.what());
	}
	catch (...) {
		result = FALSE;
		XLL_ERROR("xll_pasted: unknown exception");
	}

	return result;
}
On<xlcOnKey> xok_pasted(ON_CTRL ON_SHIFT "D", "XLL.PASTED");
