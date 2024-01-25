// excel.h - Call Excel entry point.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
// https://xlladdins.github.io/Excel4Macros/
#pragma once
#include <array>
#include <vector>
#include "oper.h"

namespace xll {

	template<class... Ts>
	inline OPER Excel(int fn, Ts... ts)
	{
		OPER res;

		std::array os{ std::move(OPER(ts))... };
		//std::vector<OPER> os{ OPER(ts)... };
		LPXLOPER12 as[sizeof...(ts)];
		for (size_t i = 0; i < os.size(); ++i) {
			as[i] = &os[i];
		}

		int ret = ::Excel12v(fn, &res, sizeof...(ts), &as[0]);
		ensure_ret(ret);
		ensure_err(res);
		if (res.xltype & xltypeAlloc) {
			res.xltype |= xlbitXLFree;
		}

		return res;
	}
	
	inline OPER Excel(int fn)
	{
		OPER res;

		const int ret = ::Excel12v(fn, &res, 0, nullptr);
		ensure_ret(ret);
		ensure_err(res);
		if (res.xltype & xltypeAlloc) {
			res.xltype |= xlbitXLFree;
		}

		return res;
	}
	
} // namespace xll
