// defines.h - Base definitions for the xll add-in library.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <span>
#include <string_view>
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>
//#include <minwindef.h>
extern "C" {
#include "XLCALL.H" // NOLINT
}
#include "ensure.h"

namespace xll {

	constexpr int xlbitFree = xlbitXLFree | xlbitDLLFree;
	constexpr WORD type(const XLOPER& x) noexcept
	{
		return x.xltype & ~(xlbitFree);
	}
	constexpr DWORD type(const XLOPER12& x)
	{
		return x.xltype & (~xlbitFree);
	}

	// xltypeX, XLOPER12::val.X, X, description
#define XLL_TYPE_SCALAR(X) \
    X(Num,  num,      double,  "IEEE 64-bit floating point") \
    X(Bool, xbool,    BOOL,    "Boolean value")              \
    X(Err,  err,      int,     "Error type")                 \
    X(Int,  w,        int,     "32-bit signed integer")      \

#define XLL_SCALAR(a, b, c, d)  | xltype##a
	constexpr int xltypeScalar = 0 
		XLL_TYPE_SCALAR(XLL_SCALAR);
#undef XLL_SCALAR
	constexpr bool isScalar(const XLOPER12& x)
	{
		return (x.xltype & xltypeScalar) != 0;
	}

	// Create XLOPER12 from scalar.
	// constexpr XLOPER12 Num(double x) { XLOPER12 o; o.xltype = xltypeNum; o.val.num = x; return o; }
#define XLL_SCALAR(a, b, c, d) constexpr XLOPER12 a(c x) { XLOPER12 o; o.xltype = xltype##a; o.val.b = x; return o; }
	XLL_TYPE_SCALAR(XLL_SCALAR)
#undef XLL_SCALAR
#ifdef _DEBUG
	static_assert(Num(1.23).xltype == xltypeNum);
	static_assert(Num(1.23).val.num == 1.23);
	static_assert(Bool(true).xltype == xltypeBool);
	static_assert(Bool(true).val.xbool == TRUE);
	static_assert(Bool(false).val.xbool == FALSE);
	static_assert(Err(xlerrNA).xltype == xltypeErr);
	static_assert(Err(xlerrNA).val.err == xlerrNA);
	//static_assert(SRef(XLREF12{ 1,2,3,4 }).xltype == xltypeSRef);
	static_assert(Int(123).xltype == xltypeInt);
	static_assert(Int(123).val.w == 123);
#endif // _DEBUG

	// Return scalar from XLOPER12.
	// constexpr double Num(const XLOPER12& x) { return x.val.num; }
#define XLL_SCALAR(a, b, c, d) constexpr c a(const XLOPER12& x) { return x.val.b; }
	XLL_TYPE_SCALAR(XLL_SCALAR)
#undef XLL_SCALAR
#ifdef _DEBUG
	static_assert(Num(Num(1.23)) == 1.23);
	static_assert(Num(Num(Num(1.23))).xltype == xltypeNum);
	static_assert(Num(Num(Num(1.23))).val.num == 1.23);
	static_assert(Bool(Bool(true)) == TRUE);
	static_assert(Err(Err(xlerrNA)) == xlerrNA);
	static_assert(Int(Int(123)) == 123);
#endif // _DEBUG

	// isNum, ...
#define XLL_IS(a, b, c, d) constexpr bool is##a(const XLOPER12& x) { return type(x) == xltype##a; }
	XLL_TYPE_SCALAR(XLL_IS)
#undef XLL_IS
#ifdef _DEBUG
	static_assert(isNum(Num(1.23)));
	static_assert(isBool(Bool(true)));
	static_assert(isErr(Err(xlerrNA)));
	static_assert(isInt(Int(123)));
#endif // _DEBUG

	// Convert to number.
	constexpr double asNum(const XLOPER12& x) noexcept
	{
		switch (type(x)) {
		case xltypeNum:
			return x.val.num;
		case xltypeBool:
			return x.val.xbool;
		case xltypeInt:
			return x.val.w;
		case xltypeMissing:
		case xltypeNil:
			return 0;
		default:
			return std::numeric_limits<double>::quiet_NaN();
		}
	}

	// Must set count to 1.
	constexpr XLOPER12 SRef(const XLREF12& ref) noexcept
	{
		XLOPER12 o;

		o.xltype = xltypeSRef;
		o.val.sref.count = 1;
		o.val.sref.ref = ref;

		return o;
	}

	// types requiring allocation where xX is pointer to data
	// xltypeX, XLOPERX::val.X, xX, description
#define XLL_TYPE_ALLOC(X) \
    X(Str,     str + 1,             XCHAR*,    "Counted string")                          \
    X(Multi,   array.lparray,       XLOPER12*, "Two dimensional array of XLOPER12 types") \
    X(Ref,     mref.lpmref->reftbl, XLREF12*,  "Array of single references")              \
    X(BigData, bigdata.h.lpbData,   BYTE*,     "Blob of binary data")                     \

// isStr, ...
#define XLL_IS(a, b, c, d) constexpr bool is##a(const XLOPER12& x) { return type(x) == xltype##a; }
	XLL_TYPE_ALLOC(XLL_IS)
#undef XLL_IS

#define XLL_ALLOC(a, b, c, d)  | xltype##a
	constexpr int xltypeAlloc = 0
		XLL_TYPE_ALLOC(XLL_ALLOC);
#undef XLL_ALLOC

	constexpr bool isAlloc(const XLOPER12& x)
	{
		return (x.xltype & xltypeAlloc) != 0;
	}

	// Return pointer to underlying data.
	// constexpr XCHAR* Str(const XLOPER12& x) { return x.xltype & xltypeStr ? x.val.str + 1 : nullptr; }
#define XLL_ALLOC(a,b,c,d)  constexpr c a(const XLOPER12& x) \
		{ return x.xltype & xltype##a ? x.val.b : nullptr; }
	XLL_TYPE_ALLOC(XLL_ALLOC)
#undef XLL_ALLOC

	// Argument for std::span(ptr, count).
	constexpr int count(const XLOPER12& x)
	{
		if (x.xltype & xltypeStr)
			return x.val.str[0];
		if (x.xltype & xltypeMulti)
			return x.val.array.rows * x.val.array.columns;
		if (x.xltype & xltypeRef)
			return x.val.mref.lpmref->count;
		if (x.xltype & xltypeBigData)
			return x.val.bigdata.cbData;

		return 0;
	}

	// xlbitX, description
#define XLL_BIT(X) \
	X(XLFree,  "Excel owns memory")    \
	X(DLLFree, "AutoFree owns memory") \

#define XLL_NULL_TYPE(X)                    \
	X(Missing, "missing function argument") \
	X(Nil,     "empty cell")                \

// isMissing, ...
#define XLL_IS(a, b) constexpr bool is##a(const XLOPER12& x) { return type(x) == xltype##a; }
	XLL_NULL_TYPE(XLL_IS)
#undef XLL_IS

#define XLL_NULL(t, d) constexpr XLOPER12 t = XLOPER12{ .xltype = xltype##t };
	XLL_NULL_TYPE(XLL_NULL)
#undef XLL_NULL

	// https://learn.microsoft.com/en-us/office/client-developer/excel/excel-worksheet-and-expression-evaluation#returning-errors
	// xlerrX, Excel error string, error description
#define XLL_TYPE_ERR(X)                                                     \
	X(Null,  "#NULL!",  "intersection of two ranges that do not intersect") \
	X(Div0,  "#DIV/0!", "formula divides by zero")                          \
	X(Value, "#VALUE!", "variable in formula has wrong type")               \
	X(Ref,   "#REF!",   "formula contains an invalid cell reference")       \
	X(Name,  "#NAME?",  "unrecognized formula name or text")                \
	X(Num,   "#NUM!",   "invalid number")                                   \
	X(NA,    "#N/A",    "value not available to a formula.")                \

#define XLL_ERR(a, b, c) constexpr XLOPER12 Err##a \
		= XLOPER12{ .val = {.err = xlerr##a}, .xltype = xltypeErr };
	XLL_TYPE_ERR(XLL_ERR)
#undef XLL_ERR
#ifdef _DEBUG
#define XLL_ERR(a, b, c) \
	static_assert(Err##a.val.err == xlerr##a); \
	static_assert(Err##a.xltype == xltypeErr);
	XLL_TYPE_ERR(XLL_ERR)
#undef XLL_ERR
#endif // _DEBUG

	// Disambiguate OPER(xlerr) constructor.
	enum class xlerr {
#define XLL_ERR_ENUM(a, b, c) a = xlerr##a, 
		XLL_TYPE_ERR(XLL_ERR_ENUM)
#undef XLL_ERR_ENUM
	};

#define XLL_ENUM(a, b, c) static_assert(static_cast<int>(xlerr::a) == xlerr##a);
	XLL_TYPE_ERR(XLL_ENUM)
#undef XLL_ENUM

	constexpr const char* xlerr_string(xlerr err)
	{
#define XLL_ERR_STRING(a, b, c) if (err == xlerr::a) return b;
		XLL_TYPE_ERR(XLL_ERR_STRING)
			return "unknown xlerr type";
#undef XLL_ERR_STRING
	}
	static_assert(std::string_view(xlerr_string(xlerr::Null)) == "#NULL!");

	constexpr const char* xlerr_description(xlerr err)
	{
#define XLL_ERR_DESC(a, b, c) if (err == xlerr::a) return b ": " c;
		XLL_TYPE_ERR(XLL_ERR_DESC)
#undef XLL_ERR_DESC
		return "unknown xlerr type";
	}
	static_assert(std::string_view(xlerr_description(xlerr::Null)) == "#NULL!: intersection of two ranges that do not intersect");

	// Check for xltypeErr.
#define ensure_err(res) ensure_message(xll::type(res) != xltypeErr, xll::xlerr_description((xll::xlerr)res.val.err));

	// https://learn.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1#data-types
	// Argument types for Excel Functions
	// XLL_XXX, Excel4, Excel12, description
#define XLL_ARG_TYPE(X)                                                          \
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

#define XLL_L(s) L##s
#define XLL_ARG(a,b,c,d) constexpr const wchar_t* XLL_##a##4 = XLL_L(b);
	XLL_ARG_TYPE(XLL_ARG)
#undef XLL_ARG

#define XLL_ARG(a,b,c,d) constexpr const wchar_t* XLL_##a = XLL_L(c);
	XLL_ARG_TYPE(XLL_ARG)
#undef XLL_ARG
#undef XLL_L

	// https://learn.microsoft.com/en-us/office/client-developer/excel/calling-into-excel-from-the-dll-or-xll#return-values 
	// xlretX, description
#define XLL_RET_TYPE(X)                                                      \
	X(xlretSuccess,                "success")                                \
	X(xlretAbort,                  "macro halted")                           \
	X(xlretInvXlfn,                "invalid function number")                \
	X(xlretInvCount,               "invalid number of arguments")            \
	X(xlretInvXloper,              "invalid OPER structure")                 \
	X(xlretStackOvfl,              "stack overflow")                         \
	X(xlretFailed,                 "command failed")                         \
	X(xlretUncalced,               "uncalculated cell")                      \
	X(xlretNotThreadSafe,          "not allowed during multi-threaded calc") \
	X(xlretInvAsynchronousContext, "invalid asynchronous function handle")   \
	X(xlretNotClusterSafe,         "not supported on cluster")               \

	inline const char* xlret_description(int ret) noexcept
	{
#define XLL_RET(a,b) if (ret == a) return b;
		XLL_RET_TYPE(XLL_RET)
#undef XLL_RET
		return "xlret type unknown";
	}
	// Check return value of Excel4/12.
	#define	ensure_ret(ret) ensure_message(ret == xlretSuccess, xll::xlret_description(ret));

} // namespace xll

// Function returning a constant value.
#define XLL_CONST(type, name, value, help, cat, topic) \
const xll::AddIn xai_ ## name (xll::Function(XLL_##type, "_xll_" #name, #name) \
.Arguments({}).FunctionHelp(help).Category(cat).HelpTopic(topic)); \
extern "C" __declspec(dllexport) type WINAPI xll_ ## name () { return value; }
