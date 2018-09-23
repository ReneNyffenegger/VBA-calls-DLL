option explicit

declare function funcsInDll lib "C:\github\VBA-calls-DLL\funcsInDll\funcsInDll.dll" (byVal hModule as long) as string()

sub main() ' {
    dim hModule as long

  '
  ' GetModuleHandle: see https://renenyffenegger.ch/notes/development/languages/VBA/Win-API/index
  '
    hModule = GetModuleHandle("VBE7.dll")

    showFuncsInDll hModule
end sub ' }

sub showFuncsInDll(byVal hModule as long) ' {

    dim funcs() as string
    dim i       as long

    funcs = funcsInDll(hModule)

    for i = 0 to uBound(funcs)
    '   debug.print funcs(i)
        cells(i+1, 1) = funcs(i)
    next i

end sub ' }
