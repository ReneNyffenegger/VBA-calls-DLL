all: ptrToStruct.dll

ptrToStruct.dll: ptrToStruct.o
	gcc -shared $< -o $@ -Wl,--add-stdcall-alias

ptrToStruct.o: ptrToStruct.c
	gcc -c $< -o $@
