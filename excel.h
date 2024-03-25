// excel.h - Call Excel entry point.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
// https://xlladdins.github.io/Excel4Macros/
#pragma once
#include <array>
#include "oper.h"

namespace xll {

	template<class... Ts>
	inline OPER Excel(int fn, Ts&&... ts)
	{
		XLOPER12 res = { .xltype = xltypeNil };

		std::array os{std::move(OPER(ts))...};
		LPXLOPER12 as[sizeof...(ts)];
		for (size_t i = 0; i < os.size(); ++i) {
			as[i] = &os[i];
		}
		// Heap corruption if OPER address passed for res.
		int ret = ::Excel12v(fn, &res, sizeof...(ts), &as[0]);
		ensure_ret(ret);
		// ensure_err(res); // allow xltypeErr to be returned
		OPER o(res);
		if (res.xltype & xltypeAlloc) {
			::Excel12(xlFree, 0, 1, &res);
		}

		return o;
	}
	
	inline OPER Excel(int fn)
	{
		XLOPER12 res = { .xltype = xltypeNil };

		const int ret = ::Excel12v(fn, &res, 0, nullptr);
		ensure_ret(ret);
		if (res.xltype & xltypeAlloc) {
			res.xltype |= xlbitXLFree;
		}
		OPER o(res);
		if (res.xltype & xltypeAlloc) {
			::Excel12(xlFree, 0, 1, &res);
		}

		return o;
	}
	
} // namespace xll
