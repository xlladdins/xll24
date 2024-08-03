// addin.cpp - AddIn information.
#include "xll.h"

using namespace xll;

// Replace nested OPER with safe handles.
OPER make_safe(const OPER& o)
{
	if (type(o) != xltypeMulti) {
		return o;
	}

	OPER o_(o);
	for (OPER& oi : o_) {
		if (isMulti(oi)) {
			handle<OPER> h(new OPER(make_safe(oi))); // TODO: fix OPER::dealloc to check for handles.
			oi = h.get();
		}
	}

	return o_;
}

AddIn xai_addin_info(
	Function(XLL_LPOPER, L"xll_addin_info", L"ADDIN.INFO")
	.Arguments({
		Arg(XLL_LPOPER, L"Name", L"is the name or register id of the add-in."),
		})
	.Uncalced()
	.Category(L"XLL")
	.FunctionHelp(L"Return information about an add-in.")
);
LPOPER WINAPI xll_addin_info(const LPOPER pname)
{
#pragma XLLEXPORT
	static OPER info;

	try {
		const Args* pargs = AddIn::find(*pname);
		if (pargs) {
			info = make_safe(pargs->markup());
		}
		else {
			info = ErrNA;
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &info;
}