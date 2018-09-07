all: ptrToStruct.dll asAny.dll

asAny.dll: asAny.o
	gcc -shared $< -o $@ -Wl,--add-stdcall-alias

asAny.o: asAny.c
	gcc -c $< -o $@

ptrToStruct.dll: ptrToStruct.o
	gcc -shared $< -o $@ -Wl,--add-stdcall-alias

ptrToStruct.o: ptrToStruct.c
	gcc -c $< -o $@
