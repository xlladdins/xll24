﻿# xll24 library

See the [xll library](https://github.com/xlladdins/xll) for
earlier versions. Against my better judgement,
I rewrote that for the n-th time because I use this every
day and work hard to be lazy. 

There is a reason why many companies still use the ancient 
[Microsoft Excel C Software Development Kit](https://learn.microsoft.com/en-us/office/client-developer/excel/welcome-to-the-excel-software-development-kit), 
It provides the highest possible performance
for integrating C, C++, and even Fortran into Excel. 
VBA, C#, and JavaScript require data to be copied into their
world and then copied back to native Excel.
Microsoft's [Python in Excel](https://www.microsoft.com/en-us/microsoft-365/python-in-excel)
actually calls over the network to do every calculation, 
as if Python isn't slow enough already.

There is a reason why many companies don't use the ancient Microsoft Excel C SDK.
The example code and documentation are difficult to use and understand.
This library makes it easy and performant to call native code from Excel.

One reason for Python's popularity is that people who know how to call
C, C++, and even Fortran, from Python have written packages to do that.
The xll library allows you to do that directly from the most popular language
in the world, Excel.

## Install

The xll library requires 64-bit Excel on Windows and Visual Studio 2022.
Run [`setup`](setup/Release/setup.msi) to install a template project called `xll` that will
show up when you create a new project in Visual Studio.

## Use

Create a new xll project in Visual Studio.

...video...

## Add-In

An `xll` add-in is a dynamic link library, or DLL, 
that exports well-known functions.
When an xll is opened in Excel it 
[dynamically loads](https://learn.microsoft.com/en-us/windows/win32/api/libloaderapi/nf-libloaderapi-loadlibrarya) 
the xll,
[looks for](https://learn.microsoft.com/en-us/windows/win32/api/libloaderapi/nf-libloaderapi-getprocaddress),
[`xlAutoOpen`](https://learn.microsoft.com/en-us/office/client-developer/excel/xlautoopen)
and calls it. There are half a dozen 
[`xlAuto` functions](https://learn.microsoft.com/en-us/office/client-developer/excel/add-in-manager-and-xll-interface-functions)
that Excel calls to manage the lifetime of the xll. The xll library implements those for you.

To add a function to be called by when Excel calls `xlAutoXXX` create an
object of type `Auto<XXX>` and specify a function to be called.
The function takes no arguments and returns 1 to indicate success or 0 for failure.
See [`auto.h`](auto.h) for the list possible values for `XXX`.

## Excel

Everything Excel has to offer is available through the [`Excel`](excel.h) function.
The first argument is a _function number_ defined in the C SDK header file
[`XLCALL.H`](XLCALL.H)
specifying the Excel function or macro to call.
Arguments for function numbers are documented in 
[Excel4Macros](https://xlladdins.github.io/Excel4Macros/index.html).

Function numbers for functions begin with `xlf` and for macros with `xlc`.
Functions have no side effects. They return a value based only on their arguments.
Macros take no arguments and only have side effects. 
They can do anything a user can do and return 1 on success or 0 on failure.

There are exceptions to this. The primary one is 
[`xlfRegister`](https://learn.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1).
It has the side effect of registering a function or macro with Excel.
There is no need to call `xlfRegister` directly.
The [`AddIn`](addin.h) class is used to register functions and macros with Excel.

### Function

An Excel function returns a result that depends only on its arguments
and has no side effects.
To call a C or C++ function from Excel use
the [`AddIn`](addin.h) class to instantiate an object
that specifies the information Excel needs.

Here is how to register `xll_hypot` as `STD.HYPOT` in Excel.
It returns a `double` and takes two `double` arguments.
```C++
AddIn xai_hypot(
    Function(XLL_DOUBLE, "xll_hypot", "STD.HYPOT")
	.Arguments({
		Arg(XLL_DOUBLE, "x", "is a number."),
		Arg(XLL_DOUBLE, "y", "is a number."),
	})
	.Category("STD")
	.FunctionHelp("Return the length of the hypotenuse of a right triangle with sides x and y.")
	.HelpTopic("https://learn.microsoft.com/en-us/cpp/c-runtime-library/reference/hypot-hypotf-hypotl-hypot-hypotf-hypotl?view=msvc-170")
);
```
The `Function` class uses the 
[named parameter idiom](https://isocpp.org/wiki/faq/ctors#named-parameter-idiom).
The member functions `Category`, `FunctionHelp`, and `HelpTopic` are optional but people using your
handiwork will appreciate it if you supply them.

You can specify a URL in `HelpTopic` that will be opened when 
[Help on this function](https://support.microsoft.com/en-us/office/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb)
is clicked in the in the 
[Insert Function](https://support.microsoft.com/en-us/office/insert-function-74474114-7c7f-43f5-bec3-096c56e2fb13)
dialog. If you don't then it defaults to `https://google.com/search?q=xll_hypot`.

Implement `xll_hypot` by calling [`std::hypot`](https://en.cppreference.com/w/cpp/numeric/math/hypot)
from the C++ standard library.
```C++
double WINAPI xll_hypot(double x, double y)
{
#pragma XLLEXPORT
	return std::hypot(x,y);
}
```
Every function registered with Excel must be declared `WINAPI`
and exported with `#pragma XLLEXPORT` in its body.
The first version of Excel was written in [Pascal](https://dl.acm.org/doi/10.1145/155360.155378)
and the `WINAPI` calling convention
is a historical artifact of that. Unlike Unix, Windows does not make functions
visible outside of a shared library unless they are explicitly exported.
The pragma does that for you.

Keep the Excel `WINAPI` function implementations simple. 
Grab the arguments you told Excel to provide,
call your platform independent function, and return the result. 
Provide a platform independent header file and library for your C and C++ code
so it can be used on Windows, Mac OS, and Linux
to get the same results displayed in Excel. 

### Macro

An Excel macro only has side effects and can do anything a user can do. 
It takes no arguments and returns 1 on success or 0 on failure.

To register a macro specify the name of the C++ function and the name Excel will use to call it.
```C++
AddIn xai_macro(
	Macro("xll_macro", "XLL.MACRO")
);
```
Then implement it.
```
int WINAPI xll_macro(void)
{
#pragma XLLEXPORT
	Excel(xlcAlert, "你好 мир"); // UTF-8 string

	return 1;
}
```

## AddIn

The [`AddIn`](addin.h) class is constructed from [`Args`](args.h).

## `Ctrl-Shift-A`

After typing `=` and the name of a function, and optionally using `<Tab>` 
to complete the name, then pressing `Ctrl-Shift-A`
will produce the names of the arguments of the function.

I don't have access to the source code of Excel, but there is a
simple way to extend this functionality. Instead of pressing
`Ctrl-Shift-A` you can press `<Backspace>` to remove the
trailing left parenthesis and then press `<Enter>`.
You will see another funny looking number, but it is not a handle. 
It is the [register id](https://learn.microsoft.com/en-us/office/client-developer/excel/xlfregisterid)
Excel uses to keep track
user defined functions.

With your cursor in the cell, pressing `Ctrl-Shift-B` will
replace the arguments you see with `Ctrl-Shift-A` with their
default values.

Pressing `Ctrl-Shift-C` will enter the default values below the cell
and provide the function corresponding to the register id to call those.
You can change the values below the cell to provide new arguments.

Pressing `Ctrl-Shift-D` will create the names you see from `Ctrl-Shift-A`
below the cell and provide the function corresponding to the register id
with those names as arguments. The function has the `Output` style
applied and the arguments have the `Input` style applied.

