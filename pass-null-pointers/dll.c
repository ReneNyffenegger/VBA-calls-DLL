#include <windows.h>

__declspec(dllexport) int __stdcall asAny(void* pointer) {

    char buf[200];

    if (! pointer) {
     // return 0 (false) to indicate that pointer was a null pointer
        return 0;
    }
  
    wsprintfA(buf, "pointer = %p, *pointer = %s", pointer, (wchar_t*) pointer); 
    MessageBoxA(0, buf, "asAny", 0);
  
 // return -1 (true) to indicate that ptrFoo was not a null pointer
    return -1;

}
