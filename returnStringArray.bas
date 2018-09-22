option explicit

declare function stringArray lib "c:\github\VBA-calls-DLL\c\returnStringArray.dll" () as string()

sub main() ' {

    dim ary() as string
    dim i     as integer

    ary = stringArray()

    for i = lBound(ary) to uBound(ary)
        debug.print i & ": " & ary(i)
    next i

end sub ' }
