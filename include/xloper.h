// xloper.h - XLOPER12 helpers
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <iostream>
#include <string_view>
#include "ref.h"
#include "utf8.h"

namespace xll {

	// Remove xlbit flags.
	constexpr unsigned xlbitFree = xlbitXLFree | xlbitDLLFree;
	constexpr WORD type(const XLOPER& x) noexcept
	{
		return x.xltype & ~(xlbitFree);
	}
	constexpr DWORD type(const XLOPER12& x)
	{
		return x.xltype & (~xlbitFree);
	}

	constexpr int xltypeScalar = xltypeNum | xltypeErr | xltypeBool | xltypeInt;
	// Type does not use allocation.
	constexpr bool isScalar(const XLOPER12& x)
	{
		return type(x) != xltypeBigData && (type(x) & xltypeScalar) != 0; // xltypeBigData is (xltypeStr | xltypeInt)
	}

	// Type requires allocation.
	constexpr bool isAlloc(const XLOPER12& x)
	{
		switch (type(x)) {
		case xltypeStr:
		case xltypeMulti:
		case xltypeRef:
		case xltypeBigData:
			return true;
		default:
			return false;
		}
	}

	// return XLFree(x) in thread-safe functions
	// Freed by Excel when no longer needed.
	constexpr LPXLOPER12 XLFree(XLOPER12& x)
	{
		if (isAlloc(x)) {
			x.xltype |= xlbitXLFree;
		}

		return &x;
	}
	// return DLLFree(x) in thread-safe functions
	// Excel calls xlAutoFree12 when no longer needed.
	constexpr LPXLOPER12 DLLFree(XLOPER12& x)
	{
		if (isAlloc(x)) {
			x.xltype |= xlbitDLLFree;
		}

		return &x;
	}

	constexpr int rows(const XLOPER12& x) noexcept
	{
		switch (type(x)) {
		case xltypeMulti:
			return x.val.array.rows;
		case xltypeSRef:
			return rows(x.val.sref.ref);
		case xltypeMissing:
		case xltypeNil:
			return 0;
		}

		return 1;
	}
	constexpr int columns(const XLOPER12& x) noexcept
	{
		switch (type(x)) {
		case xltypeMulti:
			return x.val.array.columns;
		case xltypeSRef:
			return columns(x.val.sref.ref);
		case xltypeMissing:
		case xltypeNil:
			return 0;
		default:
			return 1;
		}
	}
	constexpr int size(const XLOPER12& x) noexcept
	{
		return rows(x) * columns(x);
	}
	// Argument for std::span(ptr, count).
	constexpr XCHAR count(const XLOPER12& x) noexcept // TODO: return int
	{
		if (type(x) == xltypeStr)
			return x.val.str[0];
		if (type(x) == xltypeMulti)
			return static_cast<XCHAR>(x.val.array.rows * x.val.array.columns);
		if (type(x) == xltypeRef)
			return static_cast<XCHAR>(x.val.mref.lpmref->count);
		if (type(x) == xltypeBigData)
			return static_cast<XCHAR>(x.val.bigdata.cbData);

		return 0;
	}

	// STL friendly
	constexpr const XLOPER12* begin(const XLOPER12& x) noexcept
	{
		return type(x) == xltypeMulti ? x.val.array.lparray : &x;
	}
	constexpr XLOPER12* begin(XLOPER12& x) noexcept
	{
		return type(x) == xltypeMulti ? x.val.array.lparray : &x;
	}
	constexpr const XLOPER12* end(const XLOPER12& x) noexcept
	{
		return type(x) == xltypeMulti ? x.val.array.lparray + size(x) : &x + 1;
	}
	constexpr XLOPER12* end(XLOPER12& x) noexcept
	{
		return type(x) == xltypeMulti ? x.val.array.lparray + size(x) : &x + 1;
	}

	// Either one row or one column Multi
	constexpr bool isVector(const XLOPER12& x) noexcept
	{
		return type(x) == xltypeMulti && (rows(x) == 1 || columns(x) == 1);
	}
	// Either two row or two column Multi
	constexpr bool isJSON(const XLOPER12& x) noexcept
	{
		return type(x) == xltypeMulti && (rows(x) == 2 || columns(x) == 2);
	}

	constexpr XLOPER12 index(const XLOPER12& x, int i) noexcept
	{
		return type(x) != xltypeMulti ? x : x.val.array.lparray[i];
	}
	constexpr XLOPER12& index(XLOPER12& x, int i) noexcept
	{
		return type(x) != xltypeMulti ? x : x.val.array.lparray[i];
	}
	constexpr XLOPER12 index(const XLOPER12& x, int i, int j) noexcept
	{
		return index(x, i * columns(x) + j);
	}
	constexpr XLOPER12& index(XLOPER12& x, int i, int j) noexcept
	{
		return index(x, i * columns(x) + j);
	}

	constexpr std::wstring_view view(const XLOPER12& x, int count)
	{
		if (type(x) == xltypeNil) {
			return std::wstring_view{};
		}

		return std::wstring_view(x.val.str + 1, count);
	}
	constexpr std::wstring_view view(const XLOPER12& x)
	{
		if (type(x) == xltypeNil) {
			return std::wstring_view{};
		}

		return std::wstring_view(x.val.str + 1, count(x));
	}

	constexpr auto span(const XLOPER12& x)
	{
		ensure(type(x) == xltypeMulti);
		
		return std::span(x.val.array.lparray, count(x));
	}

	constexpr auto ref(const XLOPER12& x)
	{
		ensure(type(x) == xltypeRef);

		return std::span(x.val.mref.lpmref->reftbl, count(x));
	}

	constexpr auto blob(const XLOPER12& x)
	{
		ensure(type(x) == xltypeBigData);

		return std::span(x.val.bigdata.h.lpbData, count(x));
	}

	// Convert XLOPER12 to double.
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

	// xltypeX, XLOPER12::val.X, X, description
#define XLL_TYPE_SCALAR(X)                               \
    X(Num,  num,   double, "IEEE 64-bit floating point") \
    X(Err,  err,   int,    "Error type")                 \

	//X(Bool, xbool, BOOL, "Boolean value")              \
	//X(Int,  w,     int,  "32-bit signed integer")      \

	// Create XLOPER12 from scalar.
	// constexpr XLOPER12 Num(double x) { XLOPER12 o; o.xltype = xltypeNum; o.val.num = x; return o; }
#define XLL_SCALAR(a, b, c, d) constexpr XLOPER12 a(c x) { XLOPER12 o; o.xltype = xltype##a; o.val.b = x; return o; }
	XLL_TYPE_SCALAR(XLL_SCALAR)
#undef XLL_SCALAR
	constexpr XLOPER12 Bool(bool b) noexcept
	{
		XLOPER12 o;
		o.xltype = xltypeBool;
		o.val.xbool = b; // TODO: ? TRUE : FALSE;
		return o;
	}
	constexpr XLOPER12 Int(int i) noexcept
	{
		XLOPER12 o;
		o.xltype = xltypeInt;
		o.val.w = i;
		return o;
	}
	constexpr XLOPER12 Long(long l) noexcept
	{
		XLOPER12 o;
		o.xltype = xltypeNum;
		o.val.num = l;
		return o;
	}
#ifdef _DEBUG
	static_assert(Num(1.23).xltype == xltypeNum);
	static_assert(Num(1.23).val.num == 1.23);
	static_assert(Bool(true).xltype == xltypeBool);
	static_assert(Bool(true).val.xbool == TRUE);
	static_assert(Bool(false).val.xbool == FALSE);
	static_assert(Err(xlerrNA).xltype == xltypeErr);
	static_assert(Err(xlerrNA).val.err == xlerrNA);
	static_assert(Int(123).xltype == xltypeInt);
	static_assert(Int(123).val.w == 123);
	static_assert(Long(123).val.num == 123);
	static_assert(Long(1234567890).val.num == 1234567890);
#pragma warning(push)
#pragma warning(disable: 4305 4309) // 'constant' : integral constant overflow
	static_assert(Long(12'345'678'901).val.num != 12345678901);
#pragma warning(pop)
#endif // _DEBUG

	// isNum, ...
#define XLL_IS(a, b, c, d) constexpr bool is##a(const XLOPER12& x) { return type(x) == xltype##a; }
	XLL_TYPE_SCALAR(XLL_IS)
#undef XLL_IS
	constexpr bool isBool(const XLOPER12& x) noexcept
	{
		return type(x) == xltypeBool;
	}
	constexpr bool isInt(const XLOPER12& x) noexcept
	{
		return type(x) == xltypeInt;
	}
	
	// Return scalar from XLOPER12.
// inline double Num(const XLOPER12& x) { return x.val.num; }
#define XLL_SCALAR(a, b, c, d) constexpr c a(const XLOPER12& x) { return x.val.b; }
	XLL_TYPE_SCALAR(XLL_SCALAR)
#undef XLL_SCALAR
	// constexpr double& Num(XLOPER12& x) { return x.val.num; }
#define XLL_SCALAR(a, b, c, d) constexpr c& a(XLOPER12& x) { return x.val.b; }
	XLL_TYPE_SCALAR(XLL_SCALAR)
#undef XLL_SCALAR
	constexpr int Int(const XLOPER12& x) noexcept
	{
		return static_cast<int>(asNum(x));
	}
	constexpr long Long(const XLOPER12& x) noexcept
	{
		return static_cast<long>(asNum(x));
	}
	// Bool defined after isTrue

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

	constexpr xlerr xlerr_array[] = {
	#define XLL_ENUM(a, b, c) xlerr::a,
		XLL_TYPE_ERR(XLL_ENUM)
	#undef XLL_ENUM
	};

	// String description of error code.
	constexpr const char* xlerr_string(xlerr err)
	{
#define XLL_ERR_STRING(a, b, c) if (err == xlerr::a) return b;
		XLL_TYPE_ERR(XLL_ERR_STRING)
#undef XLL_ERR_STRING
			return "unknown xlerr type";
	}
#ifdef _DEBUG
	static_assert(std::string_view(xlerr_string(xlerr::Null)) == "#NULL!");
#endif // _DEBUG

	constexpr const wchar_t* xlerr_wstring(xlerr err)
	{
#define XLL_ERR_STRING(a, b, c) if (err == xlerr::a) return L##b;
		XLL_TYPE_ERR(XLL_ERR_STRING)
#undef XLL_ERR_STRING
			return L"unknown xlerr type";
	}
	static_assert(std::string_view(xlerr_string(xlerr::Null)) == "#NULL!");

	constexpr const char* xlerr_description(xlerr err)
	{
#define XLL_ERR_DESC(a, b, c) if (err == xlerr::a) return b ": " c;
		XLL_TYPE_ERR(XLL_ERR_DESC)
#undef XLL_ERR_DESC
			return "unknown xlerr type";
	}
#ifdef _DEBUG
	static_assert(std::string_view(xlerr_description(xlerr::Null)) == "#NULL!: intersection of two ranges that do not intersect");
#endif // _DEBUG

	// Must set count to 1.
	constexpr XLOPER12 SRef(const XLREF12& ref) noexcept
	{
		XLOPER12 o;

		o.xltype = xltypeSRef;
		o.val.sref.count = 1;
		o.val.sref.ref = ref;

		return o;
	}
	constexpr XLREF12 SRef(const XLOPER12& sref) noexcept
	{
		return sref.val.sref.ref;
	}
	constexpr bool isSRef(const XLOPER12& x) noexcept
	{
		return type(x) == xltypeSRef;
	}

	// types requiring allocation where xX is pointer to data
	// xltypeX, XLOPERX::val.X, xX, description
#define XLL_TYPE_ALLOC(X) \
    X(Str,     str + 1,             XCHAR*,    "Pascal counted string")                          \
    X(Multi,   array.lparray,       XLOPER12*, "Two dimensional array of XLOPER12 types") \
    X(Ref,     mref.lpmref->reftbl, XLREF12*,  "Array of single references")              \
    X(BigData, bigdata.h.lpbData,   BYTE*,     "Blob of binary data")                     \

	// isStr, ...
#define XLL_IS(a, b, c, d) constexpr bool is##a(const XLOPER12& x) { return type(x) == xltype##a; }
	XLL_TYPE_ALLOC(XLL_IS)
#undef XLL_IS

	// Return pointer to underlying data.
	// constexpr XCHAR* Str(const XLOPER12& x) { return x.xltype & xltypeStr ? x.val.str + 1 : nullptr; }
#define XLL_ALLOC(a,b,c,d)  constexpr c a(const XLOPER12& x) \
		{ return type(x) == xltype##a ? x.val.b : nullptr; }
	XLL_TYPE_ALLOC(XLL_ALLOC)
#undef XLL_ALLOC

	// s must be counted Pascal string with lifetime.
	constexpr XLOPER PStr(const char* s)
	{
		return XLOPER{
			.val = {.str = const_cast<char*>(s)},
			.xltype = xltypeStr
		};
	}
	constexpr XLOPER12 PStr(const wchar_t* s)
	{
		return XLOPER12{
			.val = {.str = const_cast<wchar_t*>(s)},
			.xltype = xltypeStr
		};
	}

	// Assumes lifetime of array.
	constexpr XLOPER12 Multi(XLOPER12* array, int rows, int columns)
	{
		return XLOPER12{
			.val = {.array = {.lparray = array, .rows = rows, .columns = columns} },
			.xltype = xltypeMulti
		};
	}

	// Assumes lifetime of data
	constexpr XLOPER12 BigData(BYTE* data, long size)
	{
		return XLOPER12{
			.val = {.bigdata = {.h = {.lpbData = data }, .cbData = size} },
			.xltype = xltypeBigData
		};
	}

#define XLL_NULL_TYPE(X)                    \
	X(Missing, "missing function argument") \
	X(Nil,     "empty cell")                \

	// isMissing, isNil
#define XLL_IS(a, b) constexpr bool is##a(const XLOPER12& x) { return type(x) == xltype##a; }
	XLL_NULL_TYPE(XLL_IS)
#undef XLL_IS

	// Missing, Nil
#define XLL_NULL(t, d) constexpr XLOPER12 t = XLOPER12{ .val = { .w = 0 }, .xltype = xltype##t };
	XLL_NULL_TYPE(XLL_NULL)
#undef XLL_NULL

	// False-like value.
	constexpr bool isFalse(const XLOPER12& x)
	{
		return (isNum(x) && (Num(x) != Num(x) || Num(x) == 0))
			|| (isStr(x) && (count(x) == 0 || Str(x) == 0))
			|| (isBool(x) && x.val.xbool == false)
			|| (isRef(x) && count(x) == 0)
			|| isErr(x)
			|| (isMulti(x) && size(x) == 0)
			|| isMissing(x)
			|| isNil(x)
			|| (isSRef(x) && size(x) == 0)
			|| (isInt(x) && Int(x) == 0)
			|| (isBigData(x) && count(x) == 0)
			;
	}
	// True-like value
	constexpr bool isTrue(const XLOPER12& x)
	{
		return !isFalse(x);
	}
	constexpr bool Bool(const XLOPER12& x) noexcept
	{
		return isTrue(x);
	}
#ifdef _DEBUG
	static_assert(isFalse(Num(0)));
	static_assert(isTrue(Num(1)));
	static_assert(isFalse(Num(std::numeric_limits<double>::quiet_NaN())));
	static_assert(isFalse(PStr(L"")));
	static_assert(isFalse(PStr(L"\0x1")));
	static_assert(isFalse(Bool(false)));
	static_assert(isTrue(Bool(true)));
	static_assert(isFalse(ErrNA));
	static_assert(isFalse(XLOPER12{ .val = {.array = {.lparray = 0, .rows = 0, .columns = 0}}, .xltype = xltypeMulti }));
	static_assert(isFalse(XLOPER12{ .val = {.array = {.lparray = 0, .rows = 1, .columns = 0}}, .xltype = xltypeMulti }));
	static_assert(isFalse(XLOPER12{ .val = {.array = {.lparray = 0, .rows = 0, .columns = 1}}, .xltype = xltypeMulti }));
	static_assert(isTrue(XLOPER12{ .val = {.array = {.lparray = 0, .rows = 1, .columns = 1}}, .xltype = xltypeMulti }));
	static_assert(isFalse(Missing));
	static_assert(!isTrue(Nil));
	static_assert(isFalse(SRef(REF(0, 0, 0, 0))));
	static_assert(isFalse(SRef(REF(0, 0, 1, 0))));
	static_assert(isFalse(SRef(REF(0, 0, 0, 1))));
	static_assert(isTrue(SRef(REF(0, 0, 1, 1))));
	static_assert(isFalse(Int(0)));
	static_assert(isTrue(Int(1)));
	static_assert(isFalse(BigData(nullptr, 0)));
#endif // _DEBUG
	// Useful constants.
	constexpr XLOPER12 True = Bool(true);
	constexpr XLOPER12 False = Bool(false);

	// String of length 0.
	constexpr wchar_t EmptyStr[] = { 0 };
	constexpr XLOPER12 Empty = PStr(EmptyStr);

	constexpr std::partial_ordering compare(const XLOPER12& x, const XLOPER12& y)
	{
		const int xtype = type(x);
		const int ytype = type(y);

		if (xtype != ytype) {
			return xtype <=> ytype;
		}

		switch (xtype) {
		case xltypeNum:
			return x.val.num <=> y.val.num;
		case xltypeStr:
			return std::lexicographical_compare_three_way(view(x).begin(), view(x).end(), view(y).begin(), view(y).end());
		case xltypeBool:
			return x.val.xbool <=> y.val.xbool;
		case xltypeErr:
			return x.val.err <=> y.val.err; // ??? always false like NaN
		case xltypeMulti:
			return rows(x) != rows(y) ? rows(x) <=> rows(y)
				: columns(x) != columns(y) ? columns(x) <=> columns(y)
				: std::lexicographical_compare_three_way(span(x).begin(), span(x).end(), span(y).begin(), span(y).end(), xll::compare);
		case xltypeInt:
			return x.val.w <=> y.val.w;
		case xltypeSRef:
			return x.val.sref.ref <=> y.val.sref.ref;
		case xltypeRef:
			return x.val.mref.idSheet != y.val.mref.idSheet ? x.val.mref.idSheet <=> y.val.mref.idSheet
				: std::lexicographical_compare_three_way(ref(x).begin(), ref(x).end(), ref(y).begin(), ref(y).end());
		case xltypeBigData:
			return count(x) != count(y) ? count(x) <=> count(y)
				: std::lexicographical_compare_three_way(blob(x).begin(), blob(x).end(), blob(y).begin(), blob(y).end());
		}

		return std::partial_ordering::equivalent;
	}

	// Lookup value corresponding to key in JSON like multi.
	// If s is 2x2 then it is row oriented.
	constexpr const XLOPER12 value(const XLOPER12& key, const XLOPER12& x)
	{
		if (rows(x) != 2 && columns(x) != 2) {
			return ErrNA;
		}

		if (rows(x) == 2) {
			for (int i = 0; i < columns(x); ++i) {
				if (xll::compare(index(x, 0, i), key) == std::partial_ordering::equivalent) {
					return index(x, 1, i);
				}
			}
		}
		else {
			for (int i = 0; i < rows(x); ++i) {
				if (xll::compare(index(x, i, 0), key) == std::partial_ordering::equivalent) {
					return index(x, i, 1);
				}
			}
		}

		return ErrNA;
	}
	
	constexpr std::string String(const XLOPER12& x)
	{
		if (isFalse(x)) {
			return {};
		}
		ensure(type(x) == xltypeStr);

		const auto s = std::unique_ptr<const char>(utf8::wcstombs(x.val.str + 1, x.val.str[0]));
		return std::string(s.get() + 1, s.get()[0]);
	}
	constexpr std::wstring WString(const XLOPER12& x)
	{
		if (isFalse(x)) {
			return {};
		}
		ensure(type(x) == xltypeStr);

		return std::wstring(x.val.str + 1, x.val.str[0]);
	}

} // namespace xll

constexpr auto operator<=>(const XLOPER12& x, const XLOPER12& y)
{
	return xll::compare(x, y);
}
constexpr bool operator==(const XLOPER12& x, const XLOPER12& y)
{
	return xll::compare(x, y) == 0;
}

#ifdef _DEBUG
static_assert((xll::Num(1.23) <=> xll::Num(1.23)) == 0);
static_assert((xll::Num(1.23) <=> xll::Num(1.24)) < 0);
static_assert((xll::Num(1.25) <=> xll::Num(1.24)) > 0);
#endif // _DEBUG

// Print XLOPER12 to stream.
// TODO: Use xlfEvaluate to convert a string to an OPER.
template<class T>
	requires std::is_same<T, wchar_t>::value || std::is_same<T, char>::value
inline std::basic_ostream<T>& operator<<(std::basic_ostream<T>& os, const XLOPER12& o)
{
	if (xll::isNum(o)) {
		os << o.val.num;
	}
	else if (xll::isStr(o)) {
		os << L"\"" << xll::view(o) << L"\"";
	}
	else if (xll::isBool(o)) {
		os << (o.val.xbool ? L"TRUE" : L"FALSE");
	}
	// isRef
	else if (xll::isErr(o)) {
		os << xll::xlerr_wstring(xll::xlerr{ o.val.err });
	}
	else if (xll::isMulti(o)) {
		os << L"{";
		for (int i = 0; i < xll::rows(o); ++i) {
			if (i > 0) {
				os << L";";
			}
			for (int j = 0; j < xll::columns(o); ++j) {
				if (j > 0) {
					os << L",";
				}
				os << xll::index(o, i, j);
			}
		}
		os << L"}";
	}
	// isSRef: use xlfReftext
	else if (xll::isMissing(o) || xll::isNil(o)) {
		os << xll::Empty;
	}
	else if (xll::isInt(o)) {
		os << o.val.w;
	}
	// isBigData
	else {
		os << L"<unknown>";
	}

	return os;
}

