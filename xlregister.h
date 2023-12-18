// xlregister.h - Excel registration functions
#pragma once
#include "xloper.h"

namespace xll {

	inline void Register(const XLOPER12& addin)
	{
		XLOPER12 as[32];

		for (int i = 0; i < 32; ++i) {
			as[i] = &addin[i];
		}

		Excel12(xlfRegister, 0, 1, &addin);
	}

} // namespace xll
