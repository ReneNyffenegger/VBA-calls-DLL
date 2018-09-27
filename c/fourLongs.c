#include <windows.h>

__declspec(dllexport) void __stdcall fourLongs(long a, long b, long c, long d) {

    char buf[200];
    wsprintf(buf, "a = %d\nb = %d\nc = %d\ne = %d", a, b, c, d);
    MessageBox(0, buf, "fourLongs", 0);

}
