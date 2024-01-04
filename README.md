# Excel add-in library

There is a reason why many companies still use the ancient 
[Microsoft Excel C SDK](https://learn.microsoft.com/en-us/office/client-developer/excel/welcome-to-the-excel-software-development-kit), 
It provides the highest possible performance
for integrating C, C++, and even Fortran into Excel. 
VBA, C#, and JavaScript require data to be marshalled into their
world and then copied back to native Excel, but at least they run on your machine.
Microsoft's [Python in Excel](https://www.microsoft.com/en-us/microsoft-365/python-in-excel)
actually calls back to the mothership to do every calculation, 
as if Python isn't slow enough already.

There is a reason why many companies don't use the ancient Microsoft Excel C SDK. 
It is notoriously difficult to use. 
The xll library makes it easy to
register native code for functions and macros. 

## Install

Install `xll.msi`. This installs the library and header files.
It also installs a template for creating new projects, natvis files for debugging,
and Code Snippets for creating new functions and macros.

## Use

In Visual Studio create a new xll project...

Add code snippets for functions and macros.

## Function

An Excel function is purely functional. 
It returns a result that depends only on the function arguments.
To call a C or C++ function from Excel you must register it with Excel.

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

## Usage

After installing the NuGet package you can find Code Snippets
for a Function and Macro add-in.

## xlAuto functions

The Excel C SDK requires a number of functions that Excel uses to 
control the lifetime of the add-in. An add-in is just a DLL that.

