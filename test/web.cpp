// web.cpp - call xlfWebservice asynchronously.
// Include necessary headers
#include <chrono>
#include <thread>
#include "xll.h"

using namespace xll;

// Function to perform the computation
void WINAPI PerformComputation(OPER&& asyncHandle, double input) {
    try {
        // Simulate a time-consuming computation
        std::this_thread::sleep_for(std::chrono::seconds((int)input));
        double result = input * 2; // Example computation

        // Prepare the result
        XLOPER12 xlResult;
        xlResult.xltype = xltypeNum;
        xlResult.val.num = result;

        // Return the result to Excel
        Excel12(xlAsyncReturn, 0, 2, &asyncHandle, &xlResult);
    }
    catch (...) {
        Excel12(xlAsyncReturn, 0, 2, &asyncHandle, &ErrNA);
    }
}

// Function implementation
AddIn xai_MyAsyncFunction(
    Function(XLL_VOID, "MyAsyncFunction", "XLL.AF")
    .Arguments({
        Arg(XLL_DOUBLE, "input", "is the input value")
    })
    .Asynchronous()
    .FunctionHelp("An example asynchronous function.")
);
void WINAPI MyAsyncFunction(double input, LPOPER asyncHandle)
{
#pragma XLLEXPORT
    std::thread(PerformComputation, *asyncHandle, input).detach();
}
