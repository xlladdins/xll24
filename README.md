# Excel add-in library

There is a reason why many companies still use the ancient 
[Microsoft Excel C SDK](https://learn.microsoft.com/en-us/office/client-developer/excel/welcome-to-the-excel-software-development-kit), 
It provides the highest possible performance
for integrating C, C++, and even Fortran into Excel. 
VBA, C# VSTO, and JavaScript require data to be marshalled into their
world and then copied back to native Excel.

There is a reason why many companies don't use the ancient Microsoft Excel C SDK. 
It is notoriously difficult to use. 
The xll library makes it easy.
Just register your functions and macros by telling excel how to call them.

## Function

An Excel function is purely functional.
Every Excel function returns a result that depends only on the function arguments.
```C++
AddIn xai_function(Function(XLL_DOUBLE XLL_DOUBLE, L"?xll_function", L"XLL.FUNCTION")
	.Arg(XLL_DOUBLE, L"arg", L"is the argument")
	.Category(L"XLL")
	.FunctionHelp(L"Return the argument")
);
```
## Macro

An Excel macro only has side effects. It can do anything a user can do. 
It takes no arguments and returns 1 on success and 0 on failure.

## Register

Functions are added to Excel by _register_ing them.
You must specify the signature (return type and argument types) of the C function,
its name, and the name Excel will use to call it. 

It is a good idea to specify the Insert Function category and 
a short function help description as well.
You can also specify a URL to more extensive documentation
that will be opened when <font color=blue><u>Help on this function</u></font> is clicked in the 
[Insert Function](https://support.microsoft.com/en-us/office/insert-function-74474114-7c7f-43f5-bec3-096c56e2fb13) 
dialog.

Macros take no arguments and return 1 if it succeeds and 0 if it fails.

## Installation

NuGet, standard Installer?

## Usage

After installing the NuGet package you can find Code Snippets
for a Function and Macro add-in.

## xlAuto functions

The Excel C SDK requires a number of functions that Excel uses to 
control the lifetime of the add-in. An add-in is just a DLL that.

