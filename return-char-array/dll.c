#include <windows.h>

__declspec(dllexport) const char* __stdcall charArray() {
    return "This string originates in C";
}
