//
//   cl /LD dll.c user32.lib OleAut32.lib /Fethe.dll /link /def:dll.def
//
#include <windows.h>

typedef int (WINAPI *callbackFunc_t)(int, BSTR);

callbackFunc_t callbackFunc;

__declspec(dllexport) void __stdcall setCallbackFunction(callbackFunc_t callbackFunc_) {
     callbackFunc = callbackFunc_;
}

__declspec(dllexport) int __stdcall callCallbackFunction(int i) {

    char*  txt    ;
    int    num    ;
    int    retVal ;
    int    wslen  ;
    BSTR   txtBstr;

    if      (i == 1) { num = 11; txt = "one"  ; }
    else if (i == 2) { num = 22; txt = "two"  ; }
    else if (i == 3) { num = 33; txt = "three"; }
    else             { num =  0; txt = "?"    ; }

    wslen = MultiByteToWideChar(CP_ACP, 0, txt, strlen(txt),       0,     0); txtBstr  = SysAllocStringLen(0, wslen);
            MultiByteToWideChar(CP_ACP, 0, txt, strlen(txt), txtBstr, wslen);

    retVal  = (*callbackFunc)(num, txtBstr);

    SysFreeString(txtBstr);

    return retVal;
}
