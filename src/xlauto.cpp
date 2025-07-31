// xlauto.cpp - xlAutoXXX functions
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#include <stdexcept>
#include "auto.h"
#include "xll.h"

#define XLL_TRACE // XLL_INFORMATION(__FUNCTION__)

using namespace xll;

// Define names for all AddIn Args using __procedure as key.
AddIn xai_register(Macro(L"?xll_register", L"XLL.REGISTER").Hide());
int WINAPI xll_register(void)
{
#pragma XLLEXPORT
	try {
		for (const auto& [regid, pargs] : RegIds()) {
			OPER args = compress(pargs->asMulti());
			OPER proc = OPER(L"__") & pargs->procedure;
			Excel(xlcDefineName, proc, args);// , Nil, Nil, true);
		}
		// Turn off OnSheet.
		Excel(xlcOnSheet, Missing, Missing);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
		return FALSE;
	}
	catch (...) {
		XLL_ERROR("Unknown exception in xll_register");
		return FALSE;
	}

	return TRUE;
}


// https://learn.microsoft.com/en-us/office/client-developer/excel/xlautoopen
// Called by Excel when the xll is opened.
extern "C" int __declspec(dllexport) WINAPI
xlAutoOpen(void)
{
	XLL_TRACE;
	try {
		ensure(Auto<xll::Open>::Call());
		ensure(Auto<xll::Register>::Call());
		ensure(Auto<xll::OpenAfter>::Call());

		// Excel(xlcOnSheet, Missing, OPER(L"XLL.REGISTER"), true);
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

// https://learn.microsoft.com/en-us/office/client-developer/excel/xlautoclose
// Called by Excel when the xll is unloaded.
extern "C" int __declspec(dllexport) WINAPI
xlAutoClose(void)
{
	XLL_TRACE;
	try {
		ensure(Auto<CloseBefore>::Call());
		ensure(Auto<Close>::Call());
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

// https://learn.microsoft.com/en-us/office/client-developer/excel/xlautoadd
// Called when the user activates the xll using the Add-In Manager. 
// This function is *not* called when Excel starts up.
extern "C" int __declspec(dllexport) WINAPI
xlAutoAdd(void)
{
	XLL_TRACE;
	try {
		ensure(Auto<Add>::Call());
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

// https://learn.microsoft.com/en-us/office/client-developer/excel/xlautoremove
// Called when user deactivates the xll using the Add-In Manager. 
// This function is *not* called when an Excel session closes.
extern "C" int __declspec(dllexport) WINAPI
xlAutoRemove(void)
{
	XLL_TRACE;
	try {
		ensure(Auto<Remove>::Call());
		ensure(Auto<Unregister>::Call());
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

// https://learn.microsoft.com/en-us/office/client-developer/excel/xlautofree-xlautofree12
// Called just after a worksheet function returns an XLOPER12 with xltype&xlbitDLLFree set.
extern "C" void __declspec(dllexport) WINAPI
xlAutoFree12(LPXLOPER12 px)
{
	XLL_TRACE;
	if (px->xltype & xlbitDLLFree) {
		px->xltype &= ~xlbitDLLFree;
		static_cast<OPER*>(px)->~OPER();
	}
}

// https://learn.microsoft.com/en-us/office/client-developer/excel/xlautoregister-xlautoregister12
// Look up name and register.
extern "C" LPXLOPER12 __declspec(dllexport) WINAPI
xlAutoRegister12(const LPXLOPER12 pxName)
{
	XLL_TRACE;
	static XLOPER12 o;

	try {
		auto addin = AddIn::find(OPER(*pxName));
		o = addin ? XlfRegister(addin) : ErrValue;
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrValue;
	}
	catch (...) {
		XLL_ERROR("Unknown exception in xlAutoRegister12");

		o = ErrValue;
	}

	return &o;
}

// https://learn.microsoft.com/en-us/office/client-developer/excel/xladdinmanagerinfo-xladdinmanagerinfo12
// Called by Microsoft Excel when the Add-in Manager is invoked for the first time.
// This function is used to provide the Add-In Manager with information about your add-in.
extern "C" const XLOPER12* WINAPI 
xlAddInManagerInfo12(LPXLOPER12 pxAction)
{
	XLL_TRACE;
	static OPER xInfo;

	// Coerce to int and check if action is 1.
	if (Excel(xlCoerce, *pxAction, OPER(xltypeInt)) == 1) {
		xInfo = AddInManagerInfo();
	}
	else {
		xInfo = ErrValue;
	}
	
	return &xInfo;
}
// https://learn.microsoft.com/en-us/office/client-developer/excel/xladdinmanagerinfo-xladdinmanagerinfo12
// Called by Microsoft Excel when the Add-in Manager is invoked for the first time.
// This function is used to provide the Add-In Manager with information about your add-in.
//extern "C" LPXLOPER WINAPI
//xlAddInManagerInfo(LPXLOPER pxAction)
//{
//	XLL_TRACE;
//	// Coerce to int and check if action is 1.
//	if (Excel(xlCoerce, *pxAction, OPER(xltypeInt)) == 1) {
//		return AddInManagerInfo();
//	}
//
//	return &x;
//}
/*
xll::AddIn xai_register(xll::Macro(L"?xll_register", L"XLL.REGISTER").Hide());
int WINAPI xll_register(void)
{
#pragma XLLEXPORT
	try {
		OPER reg = Excel(xlfGetName, OPER(L"_xll_register"));
		// Define hidden name procedure poining at arguments.
		OPER key(L"_xll_register");
		//OPER key = OPER("_") & pargs->procedure;
		//OPER value = compress(pargs->asMulti());
		OPER value = OPER("foo");
		Excel(xlcDefineName, key, value);// , Nil, Nil, true);

	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
		return FALSE;
	}
	catch (...) {
		XLL_ERROR("Unknown exception in xll_register");
		return FALSE;
	}

	return TRUE;
}
*/