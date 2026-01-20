# xll24 library

See the [xll library](https://github.com/xlladdins/xll) for
an earlier version.

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
The example code and documentation are difficult to understand and use.
This library makes it easy and performant to call native code from Excel.

One reason for Python's popularity is that people who know how to call
C, C++, and even Fortran, from Python have written packages to do that.
The xll library allows you to do that directly from the most popular language
in the world, Excel.

## Install

The xll library requires 64-bit Excel on Windows and Visual Studio 2026.  
Go to <a href="https://github.com/xlladdins/xll_template>xll_template</a>
click on <button>>style="background-color: green; color: white; padding: 8px 14px; border: none; border-radius: 4px;>Use this template</a>,


Open Visual Studio 2026 and create a new Empty Project.

## Add-In

An `xll` add-in is a dynamic link library, or DLL, 
that exports well-known functions.
When an xll is opened in Excel it 
[dynamically loads](https://learn.microsoft.com/en-us/windows/win32/api/libloaderapi/nf-libloaderapi-loadlibrarya) 
the xll,
[looks for](https://learn.microsoft.com/en-us/windows/win32/api/libloaderapi/nf-libloaderapi-getprocaddress) 
[`xlAutoOpen`](https://learn.microsoft.com/en-us/office/client-developer/excel/xlautoopen), then
and calls it. There are 7 
[`xlAuto` functions](https://learn.microsoft.com/en-us/office/client-developer/excel/add-in-manager-and-xll-interface-functions)
that Excel calls to manage the lifetime of the xll. The xll library implements those for you.

To add a function to be called when Excel calls `xlAutoXXX` create an
object of type `Auto<XXX>` and specify a function to be called.
The function takes no arguments and returns 1 to indicate success or 0 for failure.
See [`auto.h`](include/auto.h) for the list possible values for `XXX`.

## Excel

Everything Excel has to offer is available through the [`Excel`](include/excel.h) function.
The first argument is a _function number_ defined in the C SDK header file
[`XLCALL.H`](include/XLCALL.H)
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
The [`AddIn`](include/addin.h) class is used to register functions and macros with Excel.

### Function

An Excel function returns a result that depends only on its arguments
and has no side effects.
To call a C or C++ function from Excel use
the [`AddIn`](include/addin.h) class to instantiate an object
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
so it can be used on any computer with a C++ compiler
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

The [`AddIn`](include/addin.h) class is constructed from [`Args`](include/args.h).
All functions and macros must be registered with Excel, and
unregistered when the xll is unloaded.

## Args

The [`Args`](include/args.h) struct is used to 
[register](https://learn.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1)
the arguments to a macro or function.
The structs `Macro` and `Function` inherit from `Args` to
to populate the `Args` struct.

Macros only require the the name of the native function and
the name Excel will use to call it.

Functions require the signature of the native function and
and allow you to provide information Excel can display to users
in the [`Insert Function`](https://support.microsoft.com/en-us/office/insert-function-74474114-7c7f-43f5-bec3-096c56e2fb13)
dialog.

The `Function` struct uses the [named parameter](https://en.wikibooks.org/wiki/More_C%2B%2B_Idioms/Named_Parameter)
idiom to specify optional arguments.

## Register

The [`XlfRegister`](register.h) function is used to register macros and functions 
with Excel.
`AddIn` objects are created when the add-in is loaded,
but there are some things that can only be done after Excel calls `xlAutoOpen`.
The `Args::prepare` function arranages the data specified in the `Args` object
into a format that is necessary to call `xlfRegister`.

## Predefined Functions and Macros

### ADDIN.INFO

The `ADDIN.INFO(name)` function returns information about an add-in
given its name.

### XLL.ALERT.LEVEL

Functions and macros can report errors, warnings, and information to the user.
The `XLL.ALERT.LEVEL` function take a mask indicating what type of alerts
to report. The mask is a sum of any of the following values:
`XLL_ALERT_ERROR()`, `XLL_ALERT_WARNING()`, and `XLL_ALERT_INFORMATION()`.

### DEPENDS

To force `cell` to be calculated after `ref`, use `=DEPENDS(cell, ref)`.

### EVAL

Selecting characters in the formula in bar then pressing `F9` evaluates
the character string. The function `EVAL(cell)` function does the same thing.

### RANGE

The `\RANGE(range)` function returns a handle to a range of cells.
The function `RANGE(handle)` returns the range corresponding to the handle.

### `Ctrl-Shift-A/B/C/D`

After typing `=` and the name of a function then pressing `Ctrl-Shift-A`
will produce the names of the arguments of the function.

I don't have access to the source code of Excel, but there is a
simple way to extend this functionality. Instead of pressing
`Ctrl-Shift-A` you can press `<Backspace>` to remove the
trailing left parenthesis and then press `<Enter>`.
You will see the [register id](https://learn.microsoft.com/en-us/office/client-developer/excel/xlfregisterid)
Excel uses to keep track of user defined functions.

With your cursor in the cell containing the register id, pressing `Ctrl-Shift-B` will
replace the arguments you see from `Ctrl-Shift-A` with their
default values.

Pressing `Ctrl-Shift-C` will enter the default values below the cell
and provide the function corresponding to the register id to call those.
You can change the values below the cell to provide new arguments.

Pressing `Ctrl-Shift-D` will define the names you see from `Ctrl-Shift-A`
below the cell and provide the function corresponding to the register id
with those names as arguments. The function has the `Output` style
applied and the arguments have the `Input` style applied.

If the function text starts with a backslash (`\`) then the
function returns a _handle_ to a C++ object. The handle is a
64-bit IEEE `double` 
having the same bits as the pointer to the underlying C++ object. 

### [Range](src/range.cpp)

The `\RANGE(cells)` function returns a handle to a range of cells.
The `RANGE(handle)` function returns the range corresponding to the handle.

## JSON

Two row `OPER`s correspond to [JSON](https://json.org) objects.
The first row are the string keys and the second row are the corresponding values.
JSON objects can be nested, and so can `OPER`s.
However, Excel only displays two dimensional ranges.
We address this by providing a handle to a range if it is nested.
Use the `RANGE` function to get the range corresponding to the handle.

## TODO

Handle multiple add-ins being loaded.
Perhaps define hidden named ranges with the data?
Only exists per Excel session.
