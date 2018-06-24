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

__declspec(dllexport) void __stdcall bits_in_long(signed long l) {

  char buf[200];
  int  i, p;

  if (sizeof(long) != 4) {
    MessageBox(0, "sizeof(long) != 4", 0, 0);
    return;
  }

  p = 32 + 6 + 3;
  buf[p+1] = 0;
  for (i = 0; i<32; i++) {

    if (i > 0 && ! (i % 8)) {
      buf[p--] = ' ';
      buf[p--] = ' ';
    }
    else if (i > 0 && ! (i % 4) ) {
      buf[p--] = ' ';
    }

    buf[p--] = (l & (1 << i)) ? 'X' : '_';

  }

  MessageBox(0, buf, "Bits in long", 0);
}
