#include <windows.h>
#include <stdio.h>

__declspec(dllexport) const char* __stdcall charArray() {
    return "This char string originates in C";
}

__declspec(dllexport) const wchar_t* __stdcall wcharArray() {
    return L"This wchar_t string originates in C";
}
