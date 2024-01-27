// auto.h - export xlAutoXXX functions
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <functional>
#include <map>
#include <vector>

// Use Auto<XXX> xao_foo(xll_foo) to run xll_foo when xlAutoXXX is called.
namespace xll {

	class Open {};
	class Register {};
	class OpenAfter {};
	class CloseBefore {};
	class Unregister {};
	class Close {};
	class Add {};
	class Remove {};

	namespace {
		using macro = std::function<int(void)>;
		inline std::vector<macro> macros;
	}

	// Register macros to be called in xlAuto functions.
	template<class T>
	struct Auto {
		Auto(macro&& m)
		{
			macros.emplace_back(m);
		}
	};

	template<class T>
	inline int Call(void)
	{
		for (const auto& m : macros) {
			if (!m()) return 0;
		}

		return 1;
	}

} // namespace xll