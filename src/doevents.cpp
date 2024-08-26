#include <Windows.h>

// https://devblogs.microsoft.com/oldnewthing/2005/02/22
void DoEvents(int ms)
{
	MSG msg;
	while (PeekMessage(&msg, nullptr, 0, 0, PM_NOREMOVE)) {
		BOOL b = GetMessage(&msg, nullptr, 0, 0);
		if (b == -1) {
			MessageBoxA(0, "GetMessage returned -1", "Error", MB_OK);
			break;
		}
		else if (b > 0) {
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
		else if (msg.message == WM_QUIT) {
#ifdef _DEBUG
			MessageBoxA(nullptr, "Calling PostQuitMessage", "Info", MB_OK);
#endif 
			PostQuitMessage(static_cast<int>(msg.wParam)); // exits when macro is done?
		}
	}
	Sleep(ms);
}
