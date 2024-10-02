// py.cpp - Generate Python ctypes module.
// Use Python[`ctypes`](https://docs.python.org/3/library/ctypes.html)
// to load and specify the signature of `xll:: AddIn` functions with .Python() to generate Python code.
// The macro `PY` generates a Python module for import.
#include <fstream>
#include <map>
#include "xll.h"

using namespace xll;

// Excel to Python types.
#define PY_ARG_TYPE(X)                      \
	X(BOOL,     A, A,  c_bool, c_bool)      \
	X(DOUBLE,   B, B,  c_double, c_double)  \
	X(CSTRING,  C, C%, c_char_p, c_wchar_p) \
	X(PSTRING,  D, D%, c_char_p, c_wchar_p) \
	X(DOUBLE_,  E, E,  c_void, c_void)      \
	X(CSTRING_, F, F%, c_void, c_void)      \
	X(PSTRING_, G, G%, c_void, c_void)      \
	X(USHORT,   H, H,  c_ushort, c_ushort)  \
	X(WORD,     H, H,  c_ushort, c_ushort)  \
	X(SHORT,    I, I,  c_short, c_short)    \
	X(LONG,     J, J,  c_long, c_long)      \
	X(FP,       K, K%, c_void_p, c_void_p)  \
	X(BOOL_,    L, L,  c_void_p, c_void_p)  \
	X(SHORT_,   M, M,  c_void_p, c_void_p)  \
	X(LONG_,    N, N,  c_void_p, c_void_p)  \
	X(LPOPER,   P, Q,  c_void_p, POINTER(OPER))    \
	X(LPXLOPER, R, U,  c_void_p, POINTER(OPER))    \
	X(UINT,     H, H,  c_uint, c_uint)     \
	X(INT,      J, J,  c_int, c_int)       \

//	X(VOLATILE, !, !,  called every time sheet is recalculated)            \
//	X(UNCALCED, #, #,  dereferencing uncalculated cells returns old value) \
//	X(VOID,     ,  >,  return type to use for asynchronous functions)      \
//	X(THREAD_SAFE,  , $, declares function to be thread safe)              \
//	X(CLUSTER_SAFE, , &, declares function to be cluster safe)             \
//	X(ASYNCHRONOUS, , X, declares function to be asynchronous)             \

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

	//
	// Excel data types.
	//

	constexpr const char* XLREF12 =
		R"(class XLREF12(Structure):
	_fields_ = [
		("rwFirst", c_int),
		("rwLast", c_int),
		("colFirst", c_int),
		("colLast", c_int),
	]
)";
	constexpr const char* SRef =
		R"(class SRef(Structure):
	_fields_ = [
		("count", c_ushort),
		("ref", XLREF12),
	]
)";

	constexpr const char* MRef =
		R"(class MRef(Structure):
	_fields_ = [
		("count", c_ushort),
		("reftbl", SRef), # only one SRef allowed
	]
)";

	constexpr const char* Array =
		R"(
class OPER(Structure):
	pass

class Array(Structure):
	_fields_ = [
		("rows", c_int),
		("cols", c_int),
		("array", POINTER(OPER)),
	]
)";

	constexpr const char* Val =
		R"(class Val(Union):
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
)"; // TODO: define bigdata

	constexpr const char* OPER =
		R"(
OPER._fields_ = [
	("val", Val),
	("xltype", c_ushort),
]
)";

	constexpr char install_root[] = R"(
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
)";

	// Replace .xll% with .py.
	inline std::wstring file_py(const xll::OPER& module)
	{
		std::wstring file = module.to_wstring();
		ensure(file.ends_with(L".xll"));
		file.replace(file.size() - 4, 4, L".py");

		return file;
	}
	// Base name of module.
	inline std::wstring module_py(const xll::OPER& module)
	{
		std::wstring file(view(module));
		ensure(file.ends_with(L".xll"));
		file.replace(file.size() - 4, 4, L"");
		auto pos = file.find_last_of(L"\\");
		file = file.substr(pos + 1);

		return file;
	}
} // namespace py

AddIn xai_make_py(Macro("xll_make_py", "PY"));
int WINAPI xll_make_py()
{
#pragma XLLEXPORT
	try {
		const xll::OPER module = Excel(xlGetName);
		const auto file = py::file_py(module);
		const auto mod = py::module_py(module);

		std::wofstream ofs;

		ofs.open(file);
		ofs << L"from ctypes import *\n";
		ofs << py::install_root;
		ofs << L"root = install_root()\n";
		ofs << L"WinDLL(root + r'\\xlcall32.dll')\n";
		ofs << mod << L" = WinDLL(r'" << view(module) << L"')\n\n";

		ofs << py::XLREF12;
		ofs << py::SRef;
		ofs << py::MRef;
		ofs << py::Array;
		ofs << py::Val;
		ofs << py::OPER;

		for (const auto& [oper, pargs] : xll::AddIns()) {
			if (pargs->is_function()) {
				OPER safe = pargs->functionText.make_safe();
				ofs << std::format(L"{} = getattr({}, r'{}')\n", view(safe), mod, view(pargs->procedure));
				auto res = view(pargs->typeText, 2);
				if (res[1] != L'%') { // not XLOPER12 type
					res = res.substr(0, 1);
				}
				ofs << std::format(L"{}.restype = {}\n", view(safe), py::ctype[res]);
				ofs << std::format(L"{}.argtypes = [", view(safe));
				for (int i = 0; i < size(pargs->argumentType); ++i) {
					if (i > 0) {
						ofs << L", ";
					}
					ofs << py::ctype[pargs->argumentType[i]];
				}
				ofs << "]\n\n";
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

