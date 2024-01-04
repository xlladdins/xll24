// excel.h - Call Excel entry point.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
// https://xlladdins.github.io/Excel4Macros/
#pragma once
#include <algorithm>
#include <array>
#include <stdexcept>
#include "oper.h"

namespace xll {

	template<class... Ts>
	inline OPER Excel(int fn, Ts... ts)
	{
		OPER res;

		std::array os{ std::move(OPER(ts))... };
		LPXLOPER12 as[sizeof...(ts)];
		std::transform(os.begin(), os.end(), as, [](OPER& o) { return &o; });

		int ret = ::Excel12v(fn, &res, sizeof...(ts), &as[0]);
		ensure_ret(ret);
		if (res.xltype & xltypeAlloc) {
			res.xltype |= xlbitXLFree;
		}

		return res;
	}

	inline OPER Excel(int fn)
	{
		OPER res;

		int ret = ::Excel12v(fn, &res, 0, nullptr);
		ensure_ret(ret);
		if (res.xltype & xltypeAlloc) {
			res.xltype |= xlbitXLFree;
		}

		return res;
	}

} // namespace xll
