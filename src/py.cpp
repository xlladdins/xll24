// py.cpp - Generate Python module using ctypes.
// https://docs.python.org/3/library/ctypes.html
// The macro `PY` generates a Python module for import.
#include <fstream>
#include <map>
#include "xll.h"

using namespace xll;

// Python preamble to call Excel add-ins.
constexpr wchar_t excel_py[] = LR"(
from ctypes import *
import winreg

def install_root():
	"""Return the root of the Excel installation on local machine."""
	try:
		key_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe'
		with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
			root, _ = winreg.QueryValueEx(key, "Path")
			return root
	except FileNotFoundError:
		return None

class XLREF12(Structure):
	"""Reference to an Excel cell range."""
	_fields_ = [
		("rwFirst", c_int),
		("rwLast", c_int),
		("colFirst", c_int),
		("colLast", c_int),
	]

class SRef(Structure):
	"""Reference to a single cell range."""
	_fields_ = [
		("count", c_ushort),
		("ref", XLREF12),
	]

class MRef(Structure):
	"""Reference to a multiple cell ranges."""
	_fields_ = [
		("count", c_ushort),
		("reftbl", SRef), # only one SRef allowed
	]

class OPER(Structure):
	"""Excel OPER data type."""
	pass

class Array(Structure):
	"""Two dimensional array of OPERs."""
	_fields_ = [
		("rows", c_int),
		("cols", c_int),
		("array", POINTER(OPER)),
	]

class Val(Union):
	_fields_ = [
		("num", c_double),
		("str", c_wchar_p),
		("xbool", c_bool),
		("err", c_int),
		("w", c_int),
		("sref", SRef),
		("mref", MRef),
		("array", Array),
		("bigdata", c_voidp),
	]

OPER._fields_ = [
	("val", Val),
	("xltype", c_ushort),
]

)";

// Excel to Python types.
#define PY_ARG_TYPE(X)                      \
	X(BOOL,     A, A,  c_bool, c_bool)      \
	X(DOUBLE,   B, B,  c_double, c_double)  \
	X(CSTRING,  C, C%, c_char_p, c_wchar_p) \
	X(PSTRING,  D, D%, c_char_p, c_wchar_p) \
	X(DOUBLE_,  E, E,  c_void_p, c_void_p)  \
	X(CSTRING_, F, F%, c_void_p, c_void_p)  \
	X(PSTRING_, G, G%, c_void_p, c_void_p)  \
	X(USHORT,   H, H,  c_ushort, c_ushort)  \
	X(WORD,     H, H,  c_ushort, c_ushort)  \
	X(SHORT,    I, I,  c_short, c_short)    \
	X(LONG,     J, J,  c_long, c_long)      \
	X(FP,       K, K%, c_void_p, c_void_p)  \
	X(BOOL_,    L, L,  c_void_p, c_void_p)  \
	X(SHORT_,   M, M,  c_void_p, c_void_p)  \
	X(LONG_,    N, N,  c_void_p, c_void_p)  \
	X(LPOPER,   P, Q,  c_void_p, POINTER(OPER)) \
	X(LPXLOPER, R, U,  c_void_p, POINTER(OPER)) \
	X(UINT,     H, H,  c_uint, c_uint)      \
	X(INT,      J, J,  c_int, c_int)        \
	X(VOID,     >, >,  None, None)          \

namespace py {

	// Excel to Python types.
	inline std::map<xll::OPER, std::wstring> ctype = {
	#define PY_ARG_CTYPE(T, A, B, C, D) { xll::OPER(L#B), L#C },
		PY_ARG_TYPE(PY_ARG_CTYPE)
	#undef PY_ARG_CTYPE
	#define PY_ARG_CTYPE(T, A, B, C, D) { xll::OPER(L#C), L#D },
		PY_ARG_TYPE(PY_ARG_CTYPE)
	#undef PY_ARG_CTYPE
	};

	// Replace .xll$ with .py.
	inline std::wstring file(const xll::OPER& name)
	{
		std::wstring file = name.to_wstring();
		ensure(file.ends_with(L".xll"));
		file.replace(file.size() - 4, 4, L".py");

		return file;
	}
	// Base name of module.
	inline std::wstring module(const xll::OPER& name)
	{
		std::wstring module(view(name));
		ensure(module.ends_with(L".xll"));
		module.replace(module.size() - 4, 4, L"");
		auto pos = module.find_last_of(L"\\");
		module = module.substr(pos + 1);

		return module;
	}
} // namespace py

// Return type and argument types for ctypes.
std::pair<std::wstring, std::vector<std::wstring>> signature(const Args* pargs)
{
	std::vector<std::wstring> args;

	auto res = py::ctype[pargs->resultType()];

	for (const auto& argi : pargs->argumentType) {
		args.push_back(py::ctype[argi]);
	}

	return { res, args };
}
// Join strings with separator.
std::wstring join(const std::vector<std::wstring>& v, const std::wstring& sep = L", ")
{
	std::wstring s;

	for (const auto& x : v) {
		if (!s.empty()) {
			s += sep;
		}
		s += x;
	}

	return s;
}

AddIn xai_py(Macro("xll_py", "PY"));
int WINAPI xll_py()
{
#pragma XLLEXPORT
	try {
		const xll::OPER name = Excel(xlGetName);
		const auto file = py::file(name);
		const auto module = py::module(name);

		std::wofstream ofs;

		ofs.open(file);
		ofs << excel_py
		    << L"root = install_root()\n"
			<< L"WinDLL(root + r'\\xlcall32.dll')\n"
			<< module << L" = WinDLL(r'" << view(name) << L"')\n\n";

		for (const auto& [_, pargs] : xll::RegIds()) {
			if (pargs->python) {
				OPER safe = pargs->functionText.safe_name();
				const auto [res, args] = signature(pargs);
				ofs << std::format(L"{} = getattr({}, r'{}')\n", view(safe), module, view(pargs->procedure))
				    << std::format(L"{}.restype = {}\n", view(safe), res)
				    << std::format(L"{}.argtypes = [{}]", view(safe), join(args));
			}
		}

		ofs.close();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
		return 0;
	}

	return 1;
}

