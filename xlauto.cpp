// xlauto.cpp - xlAutoXXX functions
#include <stdexcept>
#include "auto.h"
#include "xll.h"

using namespace xll;

// Called by Excel when the xll is opened.
extern "C" int __declspec(dllexport) WINAPI
xlAutoOpen(void)
{
	try {
		ensure (Auto<Open>::Call());
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");

		return FALSE;
	}

	return TRUE;
}

// Called by Excel when the xll is unloaded.
extern "C" int __declspec(dllexport) WINAPI
xlAutoClose(void)
{
	try {
		// UnRegister all functions and commands.
		ensure (Auto<Close>::Call());
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");

		return FALSE;
	}

	return TRUE;
}
#pragma region xlAutoAdd
// Called when the user activates the xll using the Add-In Manager. 
// This function is *not* called when Excel starts up.
extern "C" int __declspec(dllexport) WINAPI
xlAutoAdd(void)
{
	try {
		if (!Auto<Add>::Call()) {
			return FALSE;
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}
	catch (...) {
		XLL_ERROR("Unknown exception in xlAutoAdd");

		return FALSE;
	}

	return TRUE;
}
#pragma endregion xlAutoAdd

#pragma region xlAutoRemove
// Called when user deactivates the xll using the Add-In Manager. 
// This function is *not* called when an Excel session closes.
extern "C" int __declspec(dllexport) WINAPI
xlAutoRemove(void)
{
	try {
		if (!Auto<Remove>::Call()) {
			return FALSE;
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}
	catch (...) {
		XLL_ERROR("Unknown exception in xlAutoRemove");

		return FALSE;
	}

	return TRUE;
}
#pragma region xlAutoRemove

#pragma region xlAutoFree12
// Called just after a worksheet function returns an XLOPER12 with xltype&xlbitDLLFree set.
extern "C" void __declspec(dllexport) WINAPI
xlAutoFree12(LPXLOPER12 px)
{
	px = px;
	/*
	if (px->xltype & xlbitDLLFree) {
		reinterpret_cast<XOPER<XLOPER12>*>(px)->~XOPER<XLOPER12>();
	}
	*/
}
#pragma region xlAutoFree12

#if 0
// Look up name and register.
extern "C" LPXLOPER __declspec(dllexport) WINAPI
xlAutoRegister(LPXLOPER pxName)
{
	static XLOPER o;

	try {
		auto px = AddIn::Map.find(*pxName);
		ensure(px != AddIn::Map.end());
		// returns an xltypeNum or xltypeErr
		o = Register(px->second);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrNA4;
	}
	catch (...) {
		XLL_ERROR("Unknown exception in xlAutoRegister");

		o = ErrNA4;
	}

	return &o;
}

// Look up name and register.
extern "C" LPXLOPER12 __declspec(dllexport) WINAPI
xlAutoRegister12(LPXLOPER12 pxName)
{
	static XLOPER12 o;

	try {
		auto px = AddIn::Map.find(*pxName);
		ensure(px != AddIn::Map.end());
		// returns an xltypeNum or xltypeErr
		o = Register(px->second);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrNA12;
	}
	catch (...) {
		XLL_ERROR("Unknown exception in xlAutoRegister12");

		o = ErrNA12;
	}

	return &o;
}
#endif // 0

#pragma region xlAddInManagerInfo12
// Called by Microsoft Excel when the Add-in Manager is invoked for the first time.
// This function is used to provide the Add-In Manager with information about your add-in.
extern "C" LPXLOPER12 WINAPI xlAddInManagerInfo12(LPXLOPER12 pxAction)
{
	static XLOPER12 o;

	try {
		pxAction = pxAction;
		/*
		// coerce to int and check if action is 1
		OPER12 xResult_ = Excel12(xlCoerce, *pxAction, OPER12(xltypeInt));
		if (xResult_ == 1) {
			o = Excel12(xlGetName);
			// return string to use
			int b = 1;
			int e = 1;
			for (int i = 1; i <= o.val.str[0]; ++i) {
				if (o.val.str[i] == '\\') {
					b = i + 1;
				}
				else if (o.val.str[i] == '.') {
					e = i;
				}
			}
			ensure(b < e);
			o = Excel12(xlfMid, o, OPER12(b), OPER12(e - b));
		}
		else {
			o = ErrNA12;
		}
		*/
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = Err(xlerrNA);
	}
	catch (...) {
		XLL_ERROR("Unknown exception in xlAddInManagerInfo12");

		o = Err(xlerrNA);
	}

	return &o;
}
#pragma endregion xlAddInManagerInfo12