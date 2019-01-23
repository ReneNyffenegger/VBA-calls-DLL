#include <windows.h>

struct FOO {
   char* str;
   int   num;
};

__declspec(dllexport) int __stdcall passPtrFOO(struct FOO* ptrFoo) {

    char buf[200];

    if (! ptrFoo) {
  
        MessageBox(0, "ptrFoo is null", "passPtrFOO", 0);
  
     // return 0 (false) to indicate that ptrFoo was a null pointer
        return 0;
    }
  
  
    wsprintf(buf, "str = %s, num = %i", ptrFoo->str, ptrFoo->num); 
    MessageBox(0, buf, "passPtrFOO", 0);
  
 // return -1 (true) to indicate that ptrFoo was not a null pointer
    return -1;

}
