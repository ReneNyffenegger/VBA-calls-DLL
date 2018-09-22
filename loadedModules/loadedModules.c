#include <windows.h>
#include <psapi.h>

__declspec(dllexport) SAFEARRAY* __stdcall loadedModules() {

  HMODULE process = GetCurrentProcess();

#define maxNofModules 256
  HMODULE modules[maxNofModules];
  DWORD   bytesNeeded;

  if (! EnumProcessModules(
      process,
      modules, // HMODULE *lphModule,
      sizeof(modules),
     &bytesNeeded
  )) {
      MessageBox(0, "EnumProcessModules", 0, 0);
  }

  int nofModules = bytesNeeded/sizeof(HANDLE);

  SAFEARRAYBOUND dimensions[2];

  dimensions[0].lLbound = 0; dimensions[0].cElements = 2;
  dimensions[1].lLbound = 0; dimensions[1].cElements = nofModules;

  SAFEARRAY *ret;
  ret = SafeArrayCreate(
       VT_VARIANT,  // The returned array is a variant array and
       2         ,  // has two dimensions which are
       dimensions   // described in the variable dimensions
  );

  VARIANT vBaseName;
  vBaseName.vt = VT_BSTR;

  VARIANT vHandle;
  vHandle.vt = VT_I4;

  LONG putIndices[2];
  wchar_t baseName[MAX_PATH];

  for (LONG i=0; i<nofModules; i++) {

   //
   // Get the module name for a specific module handle.
   // Note: we're calling the wide character version of the
   // function (ending in W). This is necessary for the
   // following call to SysAllocString.
   //
      GetModuleBaseNameW(process, modules[i], baseName, MAX_PATH);

      vBaseName.bstrVal = SysAllocString(baseName);
      vHandle.lVal      =(long) modules[i];

   //
   // Set the array value with indices (0, i) to the
   // value of the module name.
   //
      putIndices[0] = 0;
      putIndices[1] = i;
      SafeArrayPutElement(ret, putIndices, &vBaseName);

   //
   // Set the array value with indices (1, i) to the
   // value of the module handle.
   //
      putIndices[0] = 1;
   // putIndices[1] = 1;

      SafeArrayPutElement(ret, putIndices, &vHandle);

  }

  return ret;

}
