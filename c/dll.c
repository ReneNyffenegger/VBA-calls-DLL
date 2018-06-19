//
//  gcc -c dll.c
//  gcc -shared dll.o -o the.dll -Wl,--add-stdcall-alias
//
#include <stdio.h>
#include <windows.h>

__declspec(dllexport) void __stdcall swap_double_ptr(double *a, double *b) {
  char buf[200];
  double t;

  sprintf(buf, "a = %f, b = %f", *a, *b);
  MessageBox(0, buf, "swap_double_ptr", 0);
  t  = *a;
  *a = *b;
  *b =  t;
}

__declspec(dllexport) double __stdcall diff_double(double a, double b) {
  char buf[200];

  sprintf(buf, "a = %f, b = %f", a, b);
  MessageBox(0, buf, "diff_double", 0);

  return a - b;
}

__declspec(dllexport) void __stdcall to_upper_char_ptr(char* p) {
  while (*p) {
   *p = toupper(*p);
    p++;
  }
}
