// web.cpp - call xlfWebservice asynchronously.
#include <chrono>
#include <thread>
#include "../xll.h"

using namespace xll;


int my_async(WORD ms, OPER handle)
{
	std::this_thread::sleep_for(std::chrono::milliseconds(ms));
	XLOPER12 arg = { .val = {.w = ms }, .xltype = xltypeInt };

	int ret = ::Excel12(xlAsyncReturn, 0, 2, &handle, &arg);

	return ret;
}

AddIn xai_async(
	Function(XLL_VOID, "xll_async", "XLL.ASYNC")
	.Arguments({
		Arg(XLL_WORD, "ms", "is the argument")
		})
	.Asynchronous()
	//.ThreadSafe()
);
void WINAPI xll_async(WORD ms, LPXLOPER12 handle)
{
#pragma XLLEXPORT
	{ // TODO: args not passed in correctly
		std::jthread t(my_async, ms, OPER(*handle));
		t.detach();
	}

	return;
}

int webservice(const wchar_t* url, XLOPER12 handle)
{
	//OPER result = Excel(xlfWebservice, url);
	url = url;
	XLOPER12 result = { .val = {.num = 1.23}, .xltype = xltypeNum };
	XLOPER12 Ret;
	int ret = ::Excel12(xlAsyncReturn, &Ret, 2, &handle, &result);
	
	return ret;
}

AddIn xai_webservice(
	Function(XLL_VOID, "xll_webservice", "XLL.WEBSERVICE")
	.Arguments({
		Arg(XLL_CSTRING, "url", "is a URL to call"),
		})
	.Asynchronous()
	.FunctionHelp("Call a web service.")
);
void WINAPI xll_webservice(const wchar_t* url, LPXLOPER12 phandle)
{
#pragma XLLEXPORT
	std::jthread t(webservice, url, *phandle);
	t.detach();

	return;
}