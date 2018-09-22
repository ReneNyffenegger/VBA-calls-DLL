option explicit

declare function loadedModules lib "C:\github\VBA-calls-DLL\loadedModules\loadedModules.dll" () as variant()

sub main() ' {

    dim modules()  as variant
    dim i          as integer
    dim nofModules as integer

    modules = loadedModules

  '
  ' with the optional rank parameter of uBound, it's possible
  ' to determine how many modules were returned. We set the
  ' rank parameter to two to indicate the second dimension
  '
    nofModules = uBound(modules, 2) 

    for i = 0 to nofModules
        debug.print modules(1, i) & ": " & modules(0, i)
    next i

end sub ' }
