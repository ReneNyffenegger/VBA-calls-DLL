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
    if not passPtrFOO(           f ) then
'   if not passPtrFOO( ptrFoo := f ) then
       msgBox "Unexpectedly, passPtrFOO returned false"
    end if

    if not passPtrFOOasAny(f) then
       msgBox "Unexpectedly, passPtrFOOasAny returned false"
    end if

  '
  ' Trying to pass a null pointer.
  ' It does not work if (or because) the parameter is declared with 'as any':
  '
    if not passPtrFOOasAny(0) then
       msgBox "Unexpectely, passPtrFOOasAny returned false"
    end if

  '
  ' Pass the number 0 as a longPtr which will be
  ' interpreted as a null pointer in the DLL:
  '
    if passNullPtrFOO( 0 ) then
       msgBox "Unexpectely, passNullPtrFOO returned true"
    end if

  '
  ' varPtr creates a pointer - BUT the BSTR f.str is
  ' not converted into a c string. Instead of »hello world", just
  ' an »h« is received.
  '
    if not passNullPtrFOO( varPtr(f) ) then
       msgBox "Unexpectedly, passNullPtrFOO returned false"
    end if

end sub
