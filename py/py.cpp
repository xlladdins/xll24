#include <fstream>
#include "py.h"

using namespace xll;

// Replace .xll% with .py.
inline std::string file_py(const OPER& module)
{
	std::string file = module.to_string();
	ensure(file.ends_with(".xll"));
	file.replace(file.size() - 4, 4, ".py");

	return file;
}

// Create a Python module to call names.
void make_py(const OPER& module, const OPER& names)
{
	try {
		std::string file = file_py(module);
		std::ofstream ofs;
		ofs.open(file); // ???check if exists
		std::println(ofs, "from ctypes import *");
		std::println(ofs, "foo = WinDLL('{}')", module.to_string());
		ofs.close();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
}
AddIn xai_make_py(Macro("xll_make_py", "PY"));
int WINAPI xll_make_py()
{
#pragma XLLEXPORT
	try {
		const auto& pi = AddIns();
		//ensure(pi);
		OPER module = Excel(xlGetName);
		OPER names(L"XLL.CONST");

		make_py(module, names);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
		return 0;
	}

	return 1;
}