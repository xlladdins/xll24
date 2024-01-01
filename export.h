// export.h - export known functions
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once

// Used to export undecorated function name from a dll.
// Put '#pragma XLLEXPORT' in every add-in function body.
#define XLLEXPORT comment(linker, "/export:" __FUNCDNAME__ "=" __FUNCTION__)

//#pragma comment(linker, "/include:" "_DllMain@12")
//#pragma comment(linker, "/export:" "XLCallVer@0")
#pragma comment(linker, "/export:xlAutoOpen@0=xlAutoOpen")
#pragma comment(linker, "/export:xlAutoClose@0=xlAutoClose")
#pragma comment(linker, "/export:xlAutoAdd@0=xlAutoAdd")
#pragma comment(linker, "/export:xlAutoRemove@0=xlAutoRemove")
#pragma comment(linker, "/export:xlAutoFree12@4=xlAutoFree12")
#pragma comment(linker, "/export:xlAutoRegister12@4=xlAutoRegister12")
#pragma comment(linker, "/export:xlAddInManagerInfo12@4=xlAddInManagerInfo12")
