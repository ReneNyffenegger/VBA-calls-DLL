all: ptrToStruct.dll asAny.dll returnStringArray.dll

asAny.dll: asAny.o
	gcc -shared $< -o $@ -Wl,--add-stdcall-alias

asAny.o: asAny.c
	gcc -c $< -o $@

ptrToStruct.dll: ptrToStruct.o
	gcc -shared $< -o $@ -Wl,--add-stdcall-alias

ptrToStruct.o: ptrToStruct.c
	gcc -c $< -o $@

returnStringArray.dll: returnStringArray.o
	gcc -shared $< -o $@ -Wl,--add-stdcall-alias -lOleAut32

returnStringArray.o: returnStringArray.c
	gcc -c $< -o $@
