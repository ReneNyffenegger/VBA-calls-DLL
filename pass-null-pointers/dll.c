#include <windows.h>

__declspec(dllexport) int __stdcall asAny(void* pointer) {

    char buf[200];

    if (! pointer) {
     // return 0 (false) to indicate that pointer was a null pointer
        return 0;
    }
  
    wsprintf(buf, "pointer = %d, *pointer = %s", pointer, (char*) pointer); 
    MessageBox(0, buf, "asAny", 0);
  
 // return -1 (true) to indicate that ptrFoo was not a null pointer
    return -1;

}
