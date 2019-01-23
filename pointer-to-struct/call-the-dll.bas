option explicit

type FOO
  str as string
  num as long
end type

'
' "Normal" operation: the c-struct / VBA-type FOO is passed
' as a pointer to the DLL:
'
declare function passPtrFOO                                _
        lib "c:\github\VBA-calls-DLL\pointer-to-struct\the.dll"  _
        alias "passPtrFOO"                                    (  _
                 ptrFoo as FOO                                   _
        ) as boolean

declare function passPtrFOOasAny        _
        lib "c:\github\VBA-calls-DLL\pointer-to-struct\the.dll"  _
        alias "passPtrFOO"                                    (  _
           NullPtrFoo as any                                     _
        ) as boolean

'
' "Hack": Using a longPtr along with byVal, a null pointer
' can be passed to the DLL:
'
declare function passNullPtrFOO        _
        lib "c:\github\VBA-calls-DLL\pointer-to-struct\the.dll"  _
        alias "passPtrFOO"                                    (  _
           byVal NullPtrFoo as longPtr                           _
        ) as boolean

    

sub main()

    dim f as FOO

    f.str = "hello world"
    f.num =  42

  '
  ' Pass the struct as pointer.
  ' Works ok
  '
    if passPtrFOO( ptrFoo := f ) then
       msgBox "passPtrFOO returned true"
    else
       msgBox "passPtrFOO returned false"
    end if

    if passPtrFOOasAny(f) then
       msgBox "passPtrFOOasAny returned true"
    else
       msgBox "passPtrFOOasAny returned false"
    end if

    if passPtrFOOasAny(0) then
       msgBox "passPtrFOOasAny returned true"
    else
       msgBox "passPtrFOOasAny returned false"
    end if

  '
  ' Pass the number 0 as a longPtr which will be
  ' interpreted as a null pointer in the DLL:
  '
    if passNullPtrFOO( 0 ) then
       msgBox "passNullPtrFOO returned true"
    else
       msgBox "passNullPtrFOO returned false"
    end if

  '
  ' varPtr creates a pointer - BUT the BSTR f.str is
  ' not converted into a c string. Instead of »hello world", just
  ' an »h« is received.
  '
    if passNullPtrFOO( varPtr(f) ) then
       msgBox "passNullPtrFOO returned true"
    else
       msgBox "passNullPtrFOO returned false"
    end if

end sub
