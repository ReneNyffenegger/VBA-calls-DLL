option explicit

declare function asAny                                              _
        lib "c:\github\VBA-calls-DLL\pass-null-pointers\the.dll" (  _
               byVal  pointer as any                                _
        ) as boolean

sub main()

    if asAny( nothing ) then
    '  msgBox "asAny returned true"
    else
       msgBox "asAny returned false"
    end if

'   if asAny( 0 ) then
'   '  msgBox "asAny returned true"
'   else
'      msgBox "asAny returned false"
'   end if

    if asAny( vbNullString ) then
    '  msgBox "asAny returned true"
    else
       msgBox "asAny returned false"
    end if

    dim text as string
    text = "Hello World"

    if asAny( text ) then
    '  msgBox "asAny returned true"
    else
       msgBox "asAny returned false"
    end if

end sub
