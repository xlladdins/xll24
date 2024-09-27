#include "py.h"

using namespace xll;

// Create a Python module to call names.
void make_py(const OPER& module, const OPER& names)
{
	try {
		std::string m = module.to_string();
		std::string n = name.to_string();
		std::string c = code.to_string();

		std::string s = m + "." + n + " = " + c;
		py::exec(s);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
}