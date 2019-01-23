#include <windows.h>

__declspec(dllexport) int __stdcall asAny(void* pointer) {


    if (! pointer) {
  
   // MessageBox(0, "pointer is null", "asAny", 0);
  
   // return 0 (false) to indicate that pointer was a null pointer
      return 0;
    }
  
    char buf[200];
    wsprintf(buf, "pointer = %d, *pointer = %s", pointer, (char*) pointer); 
    
    MessageBox(0, buf, "asAny", 0);
  
//  return -1 (true) to indicate that ptrFoo was not a null pointer
    return -1;

}
