// dllmain.cpp
// Copyright (c) KALX, LLC. All rights reserved. No warranty is made.
#define WIN32_LEAN_AND_MEAN
#include <Windows.h>

extern "C" BOOL WINAPI
DllMain(HINSTANCE hDLL, ULONG reason, [[maybe_unused]] LPVOID lpReserved)
{
	switch (reason) {
	case DLL_PROCESS_ATTACH:
		DisableThreadLibraryCalls(hDLL);
		break;
	case DLL_THREAD_ATTACH:
		break;
	case DLL_THREAD_DETACH:
		break;
	case DLL_PROCESS_DETACH:
		break;
	}

	return TRUE;
}
