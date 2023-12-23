// defines.h - Base definitions for the xll add-in library.
#pragma once
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

// types requiring allocation where xX is pointer to data
// xltypeX, XLOPERX::val.X, xX, XLL_X, description
#define XLL_TYPE_ALLOC(X) \
    X(Str,     str,     str, PSTRING, "Pointer to a counted Pascal string")    \
    X(Multi,   array,   multi, LPOPER,  "Two dimensional array of OPER types") \
    X(Ref,     mref.lpmref,    lpmref, LPOPER,  "Multiple reference")          \
    X(BigData, bigdata.h.lpbData, bigdata, LPOPER,  "Blob of binary data")     \

	constexpr auto view(const XLOPER12& x) noexcept
	{
		return std::wstring_view(x.val.str + 1, x.val.str[0]);
	}
	constexpr auto span(const XLOPER12& x) noexcept
	{
		return std::span(x.val.array.lparray, x.val.array.rows * x.val.array.columns);
	}
	constexpr auto ref(const XLOPER12& x) noexcept
	{
		return std::span(x.val.mref.lpmref->reftbl, x.val.mref.lpmref->count);
	}
	constexpr auto blob(const XLOPER12& x) noexcept
	{
		return std::span(x.val.bigdata.h.lpbData, x.val.bigdata.cbData);
	}

// xlbitX, description
#define XLL_BIT(X) \
	X(XLFree,  "Excel owns memory")    \
	X(DLLFree, "AutoFree owns memory") \

	constexpr int xlbitFree = xlbitXLFree | xlbitDLLFree;

	constexpr int type(const XLOPER& x) noexcept
	{
		return x.xltype & ~(xlbitFree);
	}
	constexpr int type(const XLOPER12& x) noexcept
	{
		return x.xltype & ~(xlbitFree);
	}

#define XLL_NULL_TYPE(X)                    \
	X(Missing, "missing function argument") \
	X(Nil,     "empty cell")                \

#define XLL_NULL(t, d) constexpr XLOPER12 t() { return XLOPER12{ .xltype = xltype##t }; }
	XLL_NULL_TYPE(XLL_NULL)
#undef XLL_NULL
	static_assert(type(Missing()) == xltypeMissing);

// xlerrX, Excel error string, error description
#define XLL_ERR(X)                                                          \
	X(Null,  "#NULL!",  "intersection of two ranges that do not intersect") \
	X(Div0,  "#DIV/0!", "formula divides by zero")                          \
	X(Value, "#VALUE!", "variable in formula has wrong type")               \
	X(Ref,   "#REF!",   "formula contains an invalid cell reference")       \
	X(Name,  "#NAME?",  "unrecognized formula name or text")                \
	X(Num,   "#NUM!",   "invalid number")                                   \
	X(NA,    "#N/A",    "value not available to a formula.")                \

	enum class xlerr {
#define XLL_ERR_ENUM(e, s, d) e = xlerr##e, 
		XLL_ERR(XLL_ERR_ENUM)
#undef XLL_ERR_ENUM
	};

	constexpr const char* xlerr_string(xlerr err)
	{
#define XLL_ERR_STRING(e, s, d) if (err == xlerr::##e) return s;
		XLL_ERR(XLL_ERR_STRING)
		return "unknown xlerr type";
#undef XLL_ERR_STRING
	}
	static_assert(std::string_view(xlerr_string(xlerr::Null)) == "#NULL!");

	constexpr const char* xlerr_desription(xlerr err)
	{
#define XLL_ERR_DESC(e, s, d) if (err == xlerr::##e) return d;
		XLL_ERR(XLL_ERR_DESC)
		return "unknown xlerr type";
#undef XLL_ERR_DESC
	}
	static_assert(std::string_view(xlerr_desription(xlerr::Null)) == "intersection of two ranges that do not intersect");

} // namespace xll