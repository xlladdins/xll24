// xll.h - Excel add-in library header file
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include "export.h"
#include "alert.h"
#include "fp.h"
#include "on.h"
#include "handle.h"
#include "addin.h"
#include "excel_time.h"
#include "enum.h"

namespace xll {

	// Set or return value of xlAddInManagerInfo.
	inline LPXLOPER12 AddInManagerInfo(const XLOPER12& info = xll::Nil)
	{
		static xll::OPER x;

		if (xll::type(info) == xltypeStr) {
			x = info;
		}

		return &x;
	}

	// Use defined names for runtime information about all loaded add-ins.
	struct AddInInfo {
		// Full path to xll.
		static const OPER& GetName()
		{
			static const OPER name = Excel(xlGetName);

			return name;
		}
		static OPER Modules()
		{
			static OPER modules(L"Modules");

			return Excel(xlfGetName, modules);
		}
	};

} // namespace xll

