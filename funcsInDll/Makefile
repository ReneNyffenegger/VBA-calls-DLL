funcsInDll.dll: funcsInDll.c
	gcc -std=c99 -shared $< -o $@ -Wl,--add-stdcall-alias -lpsapi -lOleAut32
