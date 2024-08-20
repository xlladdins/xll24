// paste.cpp - Paste shortcuts.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#include "xll.h"
#include "macrofun.h"

using namespace xll;

/*
#define XLL_RGB_COLOR_ENUM(a, b, c) XLL_CONST(LONG, ENUM_COLOR_ ## a, b, "Color", "XLL Enum", "https://en.wikipedia.org/wiki/RGB_color_model");
XLL_RGB_COLOR(XLL_RGB_COLOR_ENUM)
#undef XLL_RGB_COLOR_ENUM
*/

// Define style for handles.
int red = 0;
int Handle = 0;
AddIn xai_handle_style(
	Macro(L"xll_handle_style", L"XLL.HANDLE.STYLE")
);
// Add red to color palette and define Handle style.
int WINAPI xll_handle_style()
{
#pragma XLLEXPORT
	if (!red) {
		red = EditColor(XLL_RGB_COLOR_RED);
		// Handle style.
		if (!Handle) {
			DefineStyle(L"Handle")
				.Number(L"\"0x\"#")
				.FormatFont(FormatFont(false).Size(8).Color(red))
				.Alignment(Alignment(false).Horizontal(Alignment::Horizontal::Center))
				.Alignment(Alignment(false).Vertical(Alignment::Vertical::Center));
		}
		Handle = 1;
	}

	return 1;
}
On<xlcOnSheet> xlos_handle_style("", "XLL.HANDLE.STYLE", true);

// Strip off leading '=' 
OPER Uneval(const XLOPER12& val)
{
	return isMissing(val) || isNil(val) ? OPER(L"")
		: StartsWith(val, L'=') ? OPER(view(val).substr(1))
		: isUDF(val) ? val & OPER(L"()")
		: val;
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
OPER Move(const XLOPER12& o, int r, int c, int w = 1, int h = 1)
{
	OPER sel = Excel(xlfOffset, o, r, c, w, h);
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
		OPER eval = Excel(xlfEvaluate, val);
		if (isMulti(eval)) {
			res = Reshape(ref, eval);
			Excel(xlSet, res, eval);
		}
		else {
			Excel(xlcFormula, val, res);
		}
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
		AlignHorizontalRight();
		FormatFont().Italic();

		// Format handle.
		OPER output = Reshape(Move(caller, 0, 1), Excel(xlfEvaluate, Formula(pargs)));

		// Expand caller to size of formula output.
		OPER active = Move(caller, rows(output), 0);

		OPER formula = OPER(L"=") & text & OPER(L"(");
		OPER comma(L"");
		for (int i = 0; i < pargs->count(); ++i) {
			formula &= comma;
			const OPER& name = pargs->argumentName[i];
			formula &= name;
			Set(active, name);
			AlignHorizontalRight();
			FormatFont().Bold();

			active = Move(active, 0, 1);
			if (isNil(pargs->argumentInit[i])) {
				Excel(xlcDefineName, name, active, Missing, Missing, Missing, Missing, true);
				Excel(xlcApplyStyle, L"Input");
				active = Move(active, 1, -1);
			}
			else {
				const auto ref = Set(active, pargs->argumentInit[i]);
				Excel(xlcDefineName, name, ref, Missing, Missing, Missing, Missing, true);
				Excel(xlcSelect, ref);
				Excel(xlcApplyStyle, L"Input");
				active = Move(active, rows(ref), -1);
			}

			comma = L", ";
		}
		formula &= OPER(L")");

		Excel(xlcFormula, formula, output);
		Excel(xlcSelect, output);
		Excel(xlcApplyStyle, isHandle(text) ? L"Handle" : L"Output");
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

// Get text of Listbox.
AddIn xai_list_macro(Macro(L"xll_list_macro", L"XLL.LIST.MACRO"));
int WINAPI xll_list_macro()
{
#pragma XLLEXPORT
	try {
		//OPER ret;
		//FMLAINFO fmla;
		//int res;
		//res = LPenHelper(xlGetFmlaInfo, &fmla);
		//res = LPenHelper(xlGetMouseInfo, &fmla);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
		return 0;
	}
	return 1;
}

AddIn xai_list(Macro(L"xll_list", L"XLL.LIST"));
int WINAPI xll_list()
{
#pragma XLLEXPORT
	try {
		OPER ret;
		auto ac = Excel(xlfActiveCell);
		auto sel = Excel(xlfSelection);
		auto val = Excel(xlCoerce, sel);
		auto list = Eval(val);
		const auto ac1 = Excel(xlfOffset, ac, 1, 1);
		const auto ac2 = Excel(xlfOffset, ac, 2, 2);
		ret = Excel(xlfCreateObject, 20, ac1, 0, 0, ac2, 0, 0); // Listbox
		ret = Excel(xlcObjectProperties, 1); // move and size with cell
		for (const auto& item : list) {
			ret = Excel(xlcAddListItem, item);
		}
		ret = Excel(xlcAssignToObject, L"XLL.LIST.MACRO");
		//Excel(xlcSelect, Excel(xlfOffset, ac, 3, 4));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
		return 0;
	}
	return 1;
}
On<xlcOnKey> xok_list(ON_CTRL ON_SHIFT "E", "XLL.LIST");
