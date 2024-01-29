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
		ensure(Call<xll::Open>());
		ensure(Call<xll::Register>());
		ensure(Call<xll::OpenAfter>());
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return false;
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");

		return false;
	}

	return true;
}

// https://learn.microsoft.com/en-us/office/client-developer/excel/xlautoclose
// Called by Excel when the xll is unloaded.
extern "C" int __declspec(dllexport) WINAPI
xlAutoClose(void)
{
	XLL_TRACE;
	try {
		ensure(Call<Close>());
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return false;
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");

		return false;
	}

	return true;
}

// https://learn.microsoft.com/en-us/office/client-developer/excel/xlautoadd
// Called when the user activates the xll using the Add-In Manager. 
// This function is *not* called when Excel starts up.
extern "C" int __declspec(dllexport) WINAPI
xlAutoAdd(void)
{
	XLL_TRACE;
	try {
		ensure(Call<Add>());
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return false;
	}
	catch (...) {
		XLL_ERROR("Unknown exception in xlAutoAdd");

		return false;
	}

	return true;
}

// https://learn.microsoft.com/en-us/office/client-developer/excel/xlautoremove
// Called when user deactivates the xll using the Add-In Manager. 
// This function is *not* called when an Excel session closes.
extern "C" int __declspec(dllexport) WINAPI
xlAutoRemove(void)
{
	XLL_TRACE;
	try {
		ensure(Call<Remove>());
		ensure(Call<Unregister>());
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return false;
	}
	catch (...) {
		XLL_ERROR("Unknown exception in xlAutoRemove");

		return false;
	}

	return true;
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
		const auto addin = AddIn::addins.find(OPER(*pxName));
		if (addin != AddIn::addins.end()) {
			o = XlfRegister(std::move(*addin->second));
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
extern "C" LPXLOPER12 WINAPI 
xlAddInManagerInfo12(LPXLOPER12 pxAction)
{
	XLL_TRACE;
	static XLOPER12 errValue{ Err(xlerrValue) };

	// Coerce to int and check if action is 1.
	if (Excel(xlCoerce, *pxAction, OPER(xltypeInt)) == 1) {
		return AddInManagerInfo();
	}
	
	return &errValue;
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
