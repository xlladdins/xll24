# xll library

There is a reason why many companies still use the ancient 
[Microsoft Excel C Software Development Kit](https://learn.microsoft.com/en-us/office/client-developer/excel/welcome-to-the-excel-software-development-kit), 
It provides the highest possible performance
for integrating C, C++, and even Fortran into Excel. 
VBA, C#, and JavaScript require data to be marshalled into their
world and then copied back to native Excel.
Microsoft's [Python in Excel](https://www.microsoft.com/en-us/microsoft-365/python-in-excel)
actually calls over the network to do every calculation, 
as if Python isn't slow enough already. 

There is a reason why many companies don't use the ancient Microsoft Excel C SDK. 
It is notoriously difficult to use. 
The xll library makes it easy to
register native code to call functions and run macros from Excel. 

## Install

Run [`setup`](setup/Release/setup.msi) to install a template project called `xll` that will
show up when you create a new project in Visual Studio.

## Use

In Visual Studio create a new [xll project](img/new_project.png).

...example...

## Excel

Everything Excel has to offer is available through the [`Excel`](excel.h) function.
The first argument is an `int` defined in
[`XLCALL.H`](XLCALL.H) provided by the C SDK 
specifying the Excel 
function or macro to call.
The arguments are documented in 
[Excel4Macros](https://xlladdins.github.io/Excel4Macros/index.html).

## Create

Write your own functions or call 3-rd party libraries using
the [`AddIn`](addin.h) class. 

### Function

An Excel function is a pure function. 
It returns a result that depends only on its arguments
and has no side effects.
To call a C or C++ function from Excel use
[`AddIn`](addin.h) to instantiate an object
that specifies the information Excel needs.

Here is how to register `xll_hypot` as `STD.HYPOT` in Excel.
If `x` and `y` are doubles then `xll_hypot(x, y)` returns the length of 
the hypotenuse of a right triangle with sides `x` and `y` as a double.
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
C++ does not have named arguments so the `Function` class uses the 
[named parameter idiom](https://isocpp.org/wiki/faq/ctors#named-parameter-idiom).

`.Category`, `.FunctionHelp`, and `.HelpTopic` are optional but people using your
handiwork will appreciate it if you supply them. 

You can specify a URL in `.HelpTopic` that will be opened when 
[Help on this function](https://support.microsoft.com/en-us/office/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb)
is clicked in the in the 
[Insert Function](https://support.microsoft.com/en-us/office/insert-function-74474114-7c7f-43f5-bec3-096c56e2fb13)
dialog. If you don't then it defaults to `https://google.com/search?q=xll_hypot`.

The next step is to implement `xll_hypot` by calling `std::hypot` from the C++ standard library.
```C++
double WINAPI xll_hypot(double x, double y)
{
#pragma XLLEXPORT
	return std::hypot(x,y);
}
```
Note every function registered with Excel must be declared `WINAPI`
and exported with `#pragma XLLEXPORT` in its body.
The first version of Excel was written in Pascal and the `WINAPI` calling convention
is a historical artifact of that. Unlike Unix, Windows does not make functions
visible outside of a shared library unless they are explicitly exported.

### Macro

An Excel macro only has side effects. It can do anything a user can do. 
It takes no arguments and returns 1 on success and 0 on failure.

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
	Excel(xlcAlert, "你好 мир");

	return 1;
}
```
The xll library converts utf-8 strings to wide character strings used by Excel.

