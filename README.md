# Excel add-in library

There is a reason why many companies
still use the ancient Microsoft
Excel C SDK. It provides the
highest possible performance
for integrating C, C++, and Fortran
into Excel. VBA, C# VSTO, 
and JavaScript require data
to be marshalled into their
world and then copied back 
to native Excel.

There is a reason why many companies
don't use the ancient Microsoft
Excel C SDK. It is notoriously
difficult to use. This library
makes it easy to call any
language from Excel.

## Function

Excel is purely functional.
Every function returns a
result that depends only
on the function arguments.

## Macro

A macro can do anything a user
can do. It takes no arguments
and returns only a success code.

## Register

Functions and macros are added to Excel by _register_ing them.
You must specify the name of the C function, its signature, and
the name Excel will use to call it. It is a good idea to specify
the Function Wizard category and a function help description as well.

Macros take no arguments and return a non-zero `int` 
if it succeeds and 0 if it fails.