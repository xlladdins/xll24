// xll.h - Excel add-in library header file
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include "XLCALL.h"
#include "auto.h"
#include "error.h"
#include "export.h"
#include "addin.h"
#include "fp.h"


// Set or return value of xlAddInManagerInfo.
inline LPXLOPER12 AddInManagerInfo(const XLOPER12& info = xll::Nil)
{
	static xll::OPER x;

	if (info.xltype == xltypeStr) {
		x = info;
	}

	return &x;
}

