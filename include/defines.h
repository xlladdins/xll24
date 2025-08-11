// defines.h - Base definitions for the xll add-in library.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <span>
#include <string_view>
#include "ensure.h"
#include "xloper.h"

namespace xll {

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
	X(UINT,     "H", "H",  "unsigned 2 byte int")                                \
	X(INT,      "J", "J",  "signed 4 byte int")                                  \

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
	// Check for xltypeErr.
	constexpr void ensure_err(const XLOPER12& res)
	{
		ensure_message(xll::type(res) != xltypeErr, xll::xlerr_description((xll::xlerr)res.val.err));
	}
	constexpr void ensure_ret(int ret)
	{
		ensure_message(ret == xlretSuccess, xll::xlret_description(ret));
	}

	// Register id given function text.
	inline double RegId(const XLOPER12& name)
	{
		XLOPER12 res;

		if (type(name) != xltypeStr) {
			return std::numeric_limits<double>::quiet_NaN();
		}
		int ret = Excel12(xlfEvaluate, &res, 1, &name);
		if (ret != xlretSuccess) {
			return std::numeric_limits<double>::quiet_NaN();
		}
		if (!isNum(res)) {
			return std::numeric_limits<double>::quiet_NaN();
		}

		return Num(res);
	}
} // namespace xll

// Function returning a constant value.
#define XLL_CONST(type, name, value, help, category, topic) \
const xll::AddIn xai_ ## name (xll::Function(XLL_##type, "_xll_" #name, #name) \
.Arguments({}).FunctionHelp(help).Category(category).HelpTopic(topic)); \
extern "C" __declspec(dllexport) type WINAPI xll_ ## name () { return value; }
