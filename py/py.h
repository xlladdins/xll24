// py.h - Generate module for calling Excel add-ins from Python.
#pragma once
#include "../include/xll.h"

namespace py {

// Excel to Python types.
#define PY_ARG_TYPE(X)                              \
	X(BOOL,     "A", "A",  "c_bool", "c_bool")      \
	X(DOUBLE,   "B", "B",  "c_double", "c_double")  \
	X(CSTRING,  "C", "C%", "c_char_p", "c_wchar_p") \
	X(PSTRING,  "D", "D%", "c_char_p", "c_wchar_p") \
	X(DOUBLE_,  "E", "E",  "c_void", "c_void")      \
	X(CSTRING_, "F", "F%", "c_void", "c_void")      \
	X(PSTRING_, "G", "G%", "c_void", "c_void")      \
	X(USHORT,   "H", "H",  "c_ushort", "c_ushort")  \
	X(WORD,     "H", "H",  "c_ushort", "c_ushort")  \
	X(SHORT,    "I", "I",  "c_short", "c_short")    \
	X(LONG,     "J", "J",  "c_long", "c_long")      \
	X(FP,       "K", "K%", "?", "?")                \
	X(BOOL_,    "L", "L",  "c_void_p", "c_void_p")  \
	X(SHORT_,   "M", "M",  "c_void_p", "c_void_p")  \
	X(LONG_,    "N", "N",  "c_void_p", "c_void_p")  \
	X(LPOPER,   "P", "Q",  "pointer", "to OPER struct (never a reference type)")    \
	X(LPXLOPER, "R", "U",  "pointer", "to XLOPER struct")                           \
	X(UINT,     "H", "H",  "c_uint", "c_uint")     \
	X(INT,      "J", "J",  "c_int", "c_int")       \

//	X(VOLATILE, "!", "!",  "called every time sheet is recalculated")            \
//	X(UNCALCED, "#", "#",  "dereferencing uncalculated cells returns old value") \
//	X(VOID,     "",  ">",  "return type to use for asynchronous functions")      \
//	X(THREAD_SAFE,  "", "$", "declares function to be thread safe")              \
//	X(CLUSTER_SAFE, "", "&", "declares function to be cluster safe")             \
//	X(ASYNCHRONOUS, "", "X", "declares function to be asynchronous")             \

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
		("ref", XLMREF12),
	]
)";

	constexpr const char* MRef =
R"(class MRef(Structure):
	_fields_ = [
		("count", c_ushort),
		("reftbl", SRef), # only one Sref allowed
	]
)";

	constexpr const char* Array =
R"(class Array(Structure):
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
		("bigdata", c_void),
	]
)";

	constexpr const char* OPER =
R"(class OPER(Structure):
	_fields_ = [
		("val", Val),
		("xltype", c_ushort),
	]
)";
	
	// Excel to Python types.
	inline std::map<xll::OPER, std::string> ctype = {
#define PY_ARG_CTYPE(T, A, B, C, D) { xll::OPER(A), C },
		PY_ARG_TYPE(PY_ARG_CTYPE)
#undef PY_ARG_CTYPE
#define PY_ARG_CTYPE(T, A, B, C, D) { xll::OPER(B), D },
		PY_ARG_TYPE(PY_ARG_CTYPE)
#undef PY_ARG_CTYPE
	};

} // namespace py
