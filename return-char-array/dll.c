//
// Because of SysAllocString, compile with oleAut32.lib:
//   gcc -shared dll.o -loleAut32 -o the.dll -Wl,--add-stdcall-alias
//
#include <windows.h>

__declspec(dllexport) const char* __stdcall charArray() {
    return "This char string originates in C";
}

__declspec(dllexport) const wchar_t* __stdcall wcharArray() {
    return L"This wchar_t string originates in C";
}

__declspec(dllexport) BSTR __stdcall bstr() {
    return SysAllocString(L"This bstr-string originates in C");
}

__declspec(dllexport) BSTR __stdcall bstr_c() {
    wchar_t* c = (wchar_t*) "This c-bstr-string originates in C";
    return SysAllocString(c);
}
