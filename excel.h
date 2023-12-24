// xlexcel.h - Call Excel entry point.
#pragma once
#include <algorithm>
#include <array>
#include <stdexcept>
#include "oper.h"

namespace xll {

	// Convert to array of pointers to LPXLOPER12.
	template<class... Ts>
	inline OPER Excel(int fn, Ts... ts)
	{
		OPER res;

		std::array os{ std::move(OPER(ts))... };
		LPXLOPER12 as[sizeof...(ts)];
		std::transform(os.begin(), os.end(), as, [](OPER& o) { return &o; });

		int ret = ::Excel12v(fn, &res, sizeof...(ts), &as[0]);
		if (ret != xlretSuccess) {
			throw std::runtime_error(xlret_description(ret));
		}
		if (!(res.xltype & xltypeScalar)) {
			res.xltype |= xlbitXLFree;
		}

		return res;
	}
	inline OPER Excel(int fn)
	{
		OPER res;

		ensure(xlretSuccess == ::Excel12v(fn, &res, 0, nullptr));
		if (!(res.xltype & xltypeScalar)) {
			res.xltype |= xlbitXLFree;
		}

		return res;
	}

} // namespace xll
