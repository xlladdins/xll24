// debug.cpp - dump memory leaks
// Copyright (c) KALX, LLC. All rights reserved. No warranty is made.
// Set _crtBreakAlloc = lineno to break on allocation.
#ifdef _DEBUG
#define _CRTDBG_MAP_ALLOC
//#define _CRTDBG_MAP_ALLOC_NEW 
#include <stdlib.h>
#include <crtdbg.h>

struct CrtDbg {
	CrtDbg(int flags = _CRTDBG_ALLOC_MEM_DF /* | _CRTDBG_LEAK_CHECK_DF | _CRTDBG_CHECK_ALWAYS_DF*/) noexcept
	{
		_CrtSetDbgFlag(_CrtSetDbgFlag(_CRTDBG_REPORT_FLAG) | flags);
		//_crtBreakAlloc = 2284;
	}
	// When information about a memory block is reported by one of the debug 
	// dump functions, this number is enclosed in braces, such as {36}.
	// Alternatively, set _crtBreakAlloc in the debugger.
	static long SetBreakAlloc(long breakAlloc) noexcept
	{
		return breakAlloc ? _CrtSetBreakAlloc(breakAlloc) : 0;
	}
	// Verifies that a specified memory block is in the local heap and that it 
	// has a valid debug heap block type identifier. Any of the last three
	// arguments may be NULL.
	static int IsMemoryBlock(const void* userData, unsigned int size,
		long* requestNumber, char** filename, int* linenumber) noexcept
	{
		return _CrtIsMemoryBlock(userData, size, requestNumber, filename, linenumber);
	}
	static long RequestNumber(const void* userData, unsigned int size)
	{
		long req;

		return IsMemoryBlock(userData, size, &req, 0, 0) ? req : 0;
	}
	CrtDbg(const CrtDbg&) = delete;
	CrtDbg& operator=(const CrtDbg&) = delete;
	~CrtDbg()
	{
		if (!(_CrtSetDbgFlag(_CRTDBG_REPORT_FLAG) & _CRTDBG_LEAK_CHECK_DF))
			_CrtDumpMemoryLeaks();
	}
};

// need to construct this before user segment
#pragma warning(disable: 4073)
#pragma init_seg(lib)
CrtDbg crtDbg;

#endif // _DEBUG
