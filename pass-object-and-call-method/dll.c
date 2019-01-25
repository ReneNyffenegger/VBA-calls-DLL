#include <windows.h>


//
// https://github.com/ReneNyffenegger/about-COM/tree/master/c/structs
//
#include "IDispatch_vTable.h"


typedef struct {
    IUnknown_vTable  *pVtbl;
}
IUnknown_interface;

typedef struct {
    IDispatch_vTable *pVtbl;
}
IDispatch_interface;


__declspec(dllexport) void __stdcall callMethodOnObj(IUnknown_interface **ppIUnknown) {

    IDispatch_interface    *pIDispatch;
    char                    buf[200];

    if  ( (*ppIUnknown)->pVtbl->QueryInterface(*ppIUnknown, &IID_IDispatch,  &pIDispatch) == NOERROR) {

        DISPID     dispid;
        DISPPARAMS dispParamsNoArgs = {NULL, NULL, 0, 0};
        BSTR       subName = L"theSub";

        if (pIDispatch->pVtbl->GetIDsOfNames(pIDispatch, &IID_NULL, &subName, 1, 0, &dispid) == S_OK) {
            pIDispatch->pVtbl->Invoke(pIDispatch, dispid, &IID_NULL, 0, DISPATCH_METHOD, &dispParamsNoArgs, NULL, NULL, NULL);
        }

        pIDispatch->pVtbl->Release(pIDispatch);
    }
}
