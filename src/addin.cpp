// addin.cpp - AddIn information.
#include "xll.h"
#if 0
using namespace xll;
AddIn xai_addin_info(
	Function(XLL_LPOPER, L"xll_addin_info", L"XLL.ADDIN.INFO")
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
		if (isMissing(*pname)) {
			info = OPER{};
			for (const auto& [k, v] : AddIns()) {
				info.append(k);
			}
			info.reshape(size(info), 1);
		}
		else {
			const Args* pargs = AddIn::find(*pname);
			if (pargs) {
				info = compress(pargs->Info());
			}
			else {
				info = ErrNA;
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &info;
}
#endif // 0