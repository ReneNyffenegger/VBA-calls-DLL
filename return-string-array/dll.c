#include <windows.h>

__declspec(dllexport)  SAFEARRAY* __stdcall stringArray() {

 // Specifiy the bounds of ONE dimenision of the SAFEARRAY:
    SAFEARRAYBOUND dim;

    dim.lLbound   = 0; // We want the array to start with index 0
    dim.cElements = 3; // We return 3 values

    SAFEARRAY *ret;
    ret = SafeArrayCreate(
         VT_BSTR,
         1      , // Array to be returned has one dimension …
        &dim      // … which is described in this (dim) parameter
    );

    LONG index = 0;

    SafeArrayPutElement(ret, &index, SysAllocString(L"foo")); index ++;
    SafeArrayPutElement(ret, &index, SysAllocString(L"bar")); index ++;
    SafeArrayPutElement(ret, &index, SysAllocString(L"baz"));

    return ret;
}
