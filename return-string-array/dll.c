//
//   This DLL must be linked to OleAut32:
//
//       cl /LD dll.c oleAut32.lib /Fethe.dll /link /def:dll.def
//

#include <windows.h>

__declspec(dllexport)  SAFEARRAY* __stdcall stringArray() {

 // Specifiy the bounds of ONE dimenision of the SAFEARRAY:
    SAFEARRAYBOUND dim;

    SAFEARRAY     *ret;
    LONG           index = 0;

    dim.lLbound   = 0; // We want the array to start with index 0
    dim.cElements = 3; // We return 3 values

    ret = SafeArrayCreate(
         VT_BSTR,
         1      , // Array to be returned has one dimension …
        &dim      // … which is described in this (dim) parameter
    );


    SafeArrayPutElement(ret, &index, SysAllocString(L"foo")); index ++;
    SafeArrayPutElement(ret, &index, SysAllocString(L"bar")); index ++;
    SafeArrayPutElement(ret, &index, SysAllocString(L"baz"));

    return ret;
}
