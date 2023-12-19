// xlexcel.h - Call Excel entry point.
#pragma once
#include <algorithm>
#include <array>
#include <stdexcept>
#include "ensure.h"
#include "xloper.h"

namespace xll {

	template<class... Ts>
	inline OPER12 Excel12(int fn, Ts... ts)
	{
		OPER12 res;

		std::array os{ std::move(OPER12(ts))... };
		LPXLOPER12 as[sizeof...(ts)];
		std::transform(os.begin(), os.end(), as, [](OPER12& o) { return &o; });

		ensure(xlretSuccess == ::Excel12v(fn, &res, sizeof...(ts), &as[0]));
		if (res.xltype == xltypeStr) {
			res.xltype |= xlbitXLFree;
		}

		return res;
	}

} // namespace xll
