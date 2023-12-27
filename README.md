# Excel add-in library

There is a reason why many companies still use the ancient 
[Microsoft Excel C SDK](https://learn.microsoft.com/en-us/office/client-developer/excel/welcome-to-the-excel-software-development-kit), 
It provides the highest possible performance
for integrating C, C++, and even Fortran into Excel. 
VBA, C# VSTO, and JavaScript require data to be marshalled into their
world and then copied back to native Excel.

There is a reason why many companies
don't use the ancient Microsoft
Excel C SDK. It is notoriously
difficult to use. The xll library
makes it easy.

## Function

Excel is purely functional.
Every function returns a
result that depends only
on the function arguments.

## Macro

A macro can do anything a user
can do. It takes no arguments
and returns 1 on success and 0 on failure.

## Register

Functions and macros are added to Excel by _register_ing them.
You must specify the name of the C function, its signature, and
the name Excel will use to call it. It is a good idea to specify
the Insert Function category and a short function help description as well.
You can also specify a URL to more extensive documentation
that will be opened when <font color=blue><u>Help on this function</u></font> is clicked in the 
[Insert Function](https://support.microsoft.com/en-us/office/insert-function-74474114-7c7f-43f5-bec3-096c56e2fb13) dialog.

Macros take no arguments and return a non-zero `int` 
if it succeeds and 0 if it fails.

## Installation

NuGet, standard Installer?

## Usage

After installing the NuGet package you can find Code Snippets
for a Function and Macro add-in.