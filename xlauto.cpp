// xlauto.cpp - xlAutoXXX functions
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#include <stdexcept>
#include "auto.h"
#include "xll.h"

#define XLL_TRACE // XLL_INFORMATION(__FUNCTION__)

using namespace xll;

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
xlAutoRegister12(LPXLOPER12 pxName)
{
	XLL_TRACE;
	static XLOPER12 o;

	try {
		const auto addin = AddIn::addins.find(OPER(*pxName));
		if (addin != AddIn::addins.end()) {
			o = XlfRegister(addin->second);
		}
		else {
			o = ErrValue;
		}
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
extern "C" LPXLOPER12 WINAPI xlAddInManagerInfo12(LPXLOPER12 pxAction)
{
	XLL_TRACE;
	static XLOPER12 err = Err(xlerrValue);
	
	// Coerce to int and check if action is 1.
	return Excel(xlCoerce, *pxAction, OPER(xltypeInt)) == 1 ? AddInManagerInfo() : &err;
}
