//
//  https://stackoverflow.com/a/1128453/180275
//
#include <windows.h>
#include <winnt.h>

HANDLE procHeap;

LPWSTR char2wchar(char* str) {

    LPWSTR wideString;
 //
 // Determine size of result string (in characters, not in bytes):
 //
    int nofChars = MultiByteToWideChar(
        CP_ACP, // Default Windows ANSI Code Page
        0,      // dwFlags
        str,
       -1,      // -1: Indicate that str is null terminated
        NULL,
        0       //  0: return required number of characters
      );

 //
 // Allocate memory for wide character string
 //
    wideString = HeapAlloc(procHeap, 0, nofChars * sizeof(WCHAR));

 //
 // Call MultiByteToWideChar again, this time to convert
 // char to wchar.
 //
    MultiByteToWideChar(CP_ACP, 0, str, -1, wideString, nofChars);

    return wideString;
}

__declspec(dllexport) SAFEARRAY* __stdcall funcsInDll(HMODULE hModule) {

 //
 // Alternatively, use LoadLibraryEx to get a handle to the module:
 //
 // HMODULE hModule = LoadLibraryEx("VBE7.dll", NULL, DONT_RESOLVE_DLL_REFERENCES);

    procHeap = GetProcessHeap();

    if (! hModule) {
       MessageBox(0, "! hModule", 0, 0);
    }

    if (((PIMAGE_DOS_HEADER)hModule)->e_magic != IMAGE_DOS_SIGNATURE) {
      MessageBox(0, "!= IMAGE_DOS_SIGNATURE", 0, 0);
    }

    PIMAGE_NT_HEADERS header = (PIMAGE_NT_HEADERS)((BYTE *)hModule + ((PIMAGE_DOS_HEADER)hModule)->e_lfanew);

    if (header->OptionalHeader.NumberOfRvaAndSizes <= 0) {
       MessageBox(0, "header->OptionalHeader.NumberOfRvaAndSizes <= 0", 0, 0);
    }

    PIMAGE_EXPORT_DIRECTORY exports = (PIMAGE_EXPORT_DIRECTORY)((BYTE *)hModule + header-> OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXPORT].VirtualAddress);

    if (exports->AddressOfNames == 0) {
    //  MessageBox(0, "exports->AddressOfNames == 0", 0, 0);
    }

    SAFEARRAYBOUND dimensions[1];

    dimensions[0].lLbound = 0; dimensions[0].cElements = exports->NumberOfNames;

    SAFEARRAY *ret;
    ret = SafeArrayCreate(
         VT_BSTR   ,  // The returned array is a variant array and
         1         ,  // has one dimension which is
         dimensions   // described in the variable dimensions
    );


    BYTE** names = (BYTE**)((int)hModule + exports->AddressOfNames);


    LONG putIndices[1];
    for (int i = 0; i < exports->NumberOfNames; i++) {

        putIndices[0] = i;

     //
     // We need a wide character in order to create a BSTR:
     //
        wchar_t* funcName_w = char2wchar( (BYTE *)hModule + (int)names[i] );

     //
     // Create BSTR and add to returned array
     //
        SafeArrayPutElement(ret, putIndices, SysAllocString( funcName_w ));

     //
     // Wide character string is not needed anymore
     //
        HeapFree(procHeap, 0, funcName_w);
    }
    return ret;

}
