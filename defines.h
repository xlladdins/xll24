// defines.h - Base definitions for the xll add-in library.
#pragma once
#include <span>
#include <string_view>
#include <span>
#define WIN32_LEAN_AND_MEAN
#include <Windows.h>
#include "XLCALL.H"
#include "ensure.h"

namespace xll {

	// xltypeX, XLOPER12::val.X, X, description
#define XLL_TYPE_SCALAR(X) \
    X(Num,  num,      double,  "IEEE 64-bit floating point")          \
    X(Bool, xbool,    BOOL,    "Boolean value")                       \
    X(Err,  err,      int,     "Error type")                          \
    X(SRef, sref.ref, XLREF12, "Single reference")                     \
    X(Int,  w,        int,     "32-bit signed integer")               \

	constexpr int xltypeScalar = xltypeNum | xltypeBool | xltypeErr | xltypeSRef | xltypeInt;

	template<int X> struct type_traits {};
#define XLL_TYPE(X, v, t, d) template<> struct type_traits<xltype##X> { using type = t; };
	XLL_TYPE_SCALAR(XLL_TYPE)
#undef XLL_TYPE
	static_assert(std::is_same_v<type_traits<xltypeNum>::type, double>);
	static_assert(std::is_same_v<type_traits<xltypeBool>::type, BOOL>);
	static_assert(std::is_same_v<type_traits<xltypeErr>::type, int>);
	static_assert(std::is_same_v<type_traits<xltypeSRef>::type, XLREF12>);
	static_assert(std::is_same_v<type_traits<xltypeInt>::type, int>);

#define XLL_SCALAR(n, v, t, d) constexpr XLOPER12 n(t x) { XLOPER12 o; o.xltype = xltype##n; o.val.v = x; return o; }
	XLL_TYPE_SCALAR(XLL_SCALAR)
#undef XLL_SCALAR
	static_assert(Num(1.23).xltype == xltypeNum);
	static_assert(Num(1.23).val.num == 1.23);
	static_assert(Bool(true).xltype == xltypeBool);
	static_assert(Bool(true).val.xbool == TRUE);
	static_assert(Bool(false).val.xbool == FALSE);
	static_assert(Err(xlerrNA).xltype == xltypeErr);
	static_assert(Err(xlerrNA).val.err == xlerrNA);
	static_assert(Int(123).xltype == xltypeInt);
	static_assert(Int(123).val.w == 123);

#define XLL_SCALAR(n, v, t, d) constexpr t n(const XLOPER12& x) { return x.val.v; }
	XLL_TYPE_SCALAR(XLL_SCALAR)
#undef XLL_SCALAR
	static_assert(Num(Num(1.23)) == 1.23);
	static_assert(Bool(Bool(true)) == TRUE);
	static_assert(Int(Int(123)) == 123);

	// types requiring allocation where xX is pointer to data
	// xltypeX, XLOPERX::val.X, xX, XLL_X, description
#define XLL_TYPE_ALLOC(X) \
    X(Str,     str,     str, PSTRING, "Pointer to a counted Pascal string")    \
    X(Multi,   array,   multi, LPOPER,  "Two dimensional array of OPER types") \
    X(Ref,     mref.lpmref,    lpmref, LPOPER,  "Multiple reference")          \
    X(BigData, bigdata.h.lpbData, bigdata, LPOPER,  "Blob of binary data")     \

	// xllbitX, description
#define XLL_BIT(X) \
	X(XLFree,  "Excel owns memory")    \
	X(DLLFree, "AutoFree owns memory") \

	constexpr int xlbitFree = xlbitXLFree | xlbitDLLFree;

#define XLL_NULL_TYPE(X)                    \
	X(Missing, "missing function argument") \
	X(Nil,     "empty cell")                \

#define XLL_NULL(t, d) constexpr XLOPER12 t() { return XLOPER12{ .xltype = xltype##t }; }
	XLL_NULL_TYPE(XLL_NULL)
#undef XLL_NULL
	static_assert(Missing().xltype == xltypeMissing);
	static_assert(Nil().xltype == xltypeNil);

	// xlerrX, Excel error string, error description
#define XLL_TYPE_ERR(X)                                                          \
	X(Null,  "#NULL!",  "intersection of two ranges that do not intersect") \
	X(Div0,  "#DIV/0!", "formula divides by zero")                          \
	X(Value, "#VALUE!", "variable in formula has wrong type")               \
	X(Ref,   "#REF!",   "formula contains an invalid cell reference")       \
	X(Name,  "#NAME?",  "unrecognized formula name or text")                \
	X(Num,   "#NUM!",   "invalid number")                                   \
	X(NA,    "#N/A",    "value not available to a formula.")                \

#define XLL_ERR(e, s, d) constexpr XLOPER12 Err##e = XLOPER12{ .val = {.err = xlerr##e}, .xltype = xltypeErr };
	XLL_TYPE_ERR(XLL_ERR)
#undef XLL_ERR
		static_assert(ErrNull.xltype == xltypeErr);
	static_assert(ErrNull.val.err == xlerrNull);

	enum class xlerr {
#define XLL_ERR_ENUM(e, s, d) e = xlerr##e, 
		XLL_TYPE_ERR(XLL_ERR_ENUM)
#undef XLL_ERR_ENUM
	};
	static_assert((int)xlerr::Null == xlerrNull);

	constexpr const char* xlerr_string(xlerr err)
	{
#define XLL_ERR_STRING(e, s, d) if (err == xlerr::##e) return s;
		XLL_TYPE_ERR(XLL_ERR_STRING)
			return "unknown xlerr type";
#undef XLL_ERR_STRING
	}
	static_assert(std::string_view(xlerr_string(xlerr::Null)) == "#NULL!");

	constexpr const char* xlerr_desription(xlerr err)
	{
#define XLL_ERR_DESC(e, s, d) if (err == xlerr::##e) return d;
		XLL_TYPE_ERR(XLL_ERR_DESC)
#undef XLL_ERR_DESC
			return "unknown xlerr type";
	}
	static_assert(std::string_view(xlerr_desription(xlerr::Null)) == "intersection of two ranges that do not intersect");

	// Argument types for Excel Functions
	// XLL_XXX, Excel4, Excel12, description
#define XLL_ARG_TYPE(X)                                                      \
	X(BOOL,     "A", "A",  "short int used as logical")                          \
	X(DOUBLE,   "B", "B",  "double")                                             \
	X(CSTRING,  "C", "C%", "XCHAR* to C style NULL terminated unicode string")   \
	X(PSTRING,  "D", "D%", "XCHAR* to Pascal style byte counted unicode string") \
	X(DOUBLE_,  "E", "E",  "pointer to double")                                  \
	X(CSTRING_, "F", "F%", "reference to a null terminated unicode string")      \
	X(PSTRING_, "G", "G%", "reference to a byte counted unicode string")         \
	X(USHORT,   "H", "H",  "unsigned 2 byte int")                                \
	X(WORD,     "H", "H",  "unsigned 2 byte int")                                \
	X(SHORT,    "I", "I",  "signed 2 byte int")                                  \
	X(LONG,     "J", "J",  "signed 4 byte int")                                  \
	X(FP,       "K", "K%", "pointer to struct FP")                               \
	X(BOOL_,    "L", "L",  "reference to a boolean")                             \
	X(SHORT_,   "M", "M",  "reference to signed 2 byte int")                     \
	X(LONG_,    "N", "N",  "reference to signed 4 byte int")                     \
	X(LPOPER,   "P", "Q",  "pointer to OPER struct (never a reference type)")    \
	X(LPXLOPER, "R", "U",  "pointer to XLOPER struct")                           \
	X(VOLATILE, "!", "!",  "called every time sheet is recalculated")            \
	X(UNCALCED, "#", "#",  "dereferencing uncalculated cells returns old value") \
	X(VOID,     "",  ">",  "return type to use for asynchronous functions")      \
	X(THREAD_SAFE,  "", "$", "declares function to be thread safe")              \
	X(CLUSTER_SAFE, "", "&", "declares function to be cluster safe")             \
	X(ASYNCHRONOUS, "", "X", "declares function to be asynchronous")             \

#define XLL_ARG(a,b,c,d) constexpr const wchar_t* XLL_##a##4 = L#b;
	XLL_ARG_TYPE(XLL_ARG)
#undef XLL_ARG

#define XLL_ARG(a,b,c,d) constexpr const wchar_t* XLL_##a = L#c;
		XLL_ARG_TYPE(XLL_ARG)
#undef XLL_ARG

	// xlretX, description
#define XLL_RET_TYPE(X) \
	X(xlretSuccess,                "success") \
	X(xlretAbort,                  "macro halted") \
	X(xlretInvXlfn,                "invalid function number") \
	X(xlretInvCount,               "invalid number of arguments") \
	X(xlretInvXloper,              "invalid OPER structure") \
	X(xlretStackOvfl,              "stack overflow") \
	X(xlretFailed,                 "command failed") \
	X(xlretUncalced,               "uncalculated cell") \
	X(xlretNotThreadSafe,          "not allowed during multi-threaded calc") \
	X(xlretInvAsynchronousContext, "invalid asynchronous function handle") \
	X(xlretNotClusterSafe,         "not supported on cluster") \

	inline const char* xlret_description(int ret)
	{
#define XLL_RET(a,b) if (ret == a) return b;
		XLL_RET_TYPE(XLL_RET)
#undef XLL_RET
		return "xlret type unknown";
	}
} // namespace xll