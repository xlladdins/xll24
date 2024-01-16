// xll.h - Excel add-in library header file
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include "export.h"
#include "error.h"
#include "fp.h"
#include "auto.h"
#include "addin.h"

// Set or return value of xlAddInManagerInfo.
inline LPXLOPER12 AddInManagerInfo(const XLOPER12& info = xll::Nil)
{
	static xll::OPER x;

	if (xll::type(info) == xltypeStr) {
		x = info;
	}

	return &x;
}

