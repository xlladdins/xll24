#pragma warning(disable: 4996)
#include <stdexcept>
#include "xll.h"

using namespace xll;

XLL_CONST(WORD, XLL_ALERT_ERROR, XLL_ALERT_ERROR,
	"ALERT.LEVEL flag for errors[1].", "XLL", "")
XLL_CONST(WORD, XLL_ALERT_WARNING, XLL_ALERT_WARNING,
	"ALERT.LEVEL flag for warnings[2].", "XLL", "")
XLL_CONST(WORD, XLL_ALERT_INFORMATION, XLL_ALERT_INFORMATION,
	"ALERT.LEVEL flag for information[4].", "XLL", "")

AddIn xai_alert_level(
	Function(XLL_WORD, "xll_alert_level_", "XLL.ALERT.LEVEL")
	.Arguments({
		Arg(XLL_LPOPER, "level", "is the alert level mask to set.", "=XLL_ALERT_ERROR()"),
	})
	.FunctionHelp("Set the current alert level using a mask and return the old mask.")
	.Category("XLL")
	.Documentation(R"(
The xll library can report errors, warnings, and information using pop-up alerts.
These can be turned on or off using any sum of <code>XLL_ALERT_ERROR()</code>,
<code>XLL_ALERT_WARNING()</code>, or <code>XLL_ALERT_INFO()</code> flags.
The function returns the previous mask and the argument
is stored at <code>HKEY_CURRENT_USER\Software\KALX\xll\xll_alert_level</code>
in the registry to persist across Excel sessions.
)")
);
DWORD WINAPI xll_alert_level_(LPOPER plevel)
{
#pragma XLLEXPORT
	DWORD oal = get_alert_mask();

	if (isNum(*plevel)) {
		set_alert_mask(static_cast<DWORD>(asNum(*plevel)));
	}

	return oal;
}
