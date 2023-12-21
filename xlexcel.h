// xlexcel.h - Call Excel entry point.
#pragma once
#include <algorithm>
#include <array>
#include <stdexcept>
#include "ensure.h"
#include "xloper.h"

namespace xll {

	// Convert to array of pointers to LPXLOPER12.
	template<class... Ts>
	inline OPER Excel(int fn, Ts... ts)
	{
		OPER res;

		std::array os{ std::move(OPER(ts))... };
		LPXLOPER12 as[sizeof...(ts)];
		std::transform(os.begin(), os.end(), as, [](OPER& o) { return &o; });

		ensure(xlretSuccess == ::Excel12v(fn, &res, sizeof...(ts), &as[0]));
		if (!(res.xltype & xltypeScalar)) {
			res.xltype |= xlbitXLFree;
		}

		return res;
	}

} // namespace xll
