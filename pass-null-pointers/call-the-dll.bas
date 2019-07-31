option explicit

declare ptrSafe function asAny        _
        lib "the.dll" (               _
               byVal  pointer as any  _
        ) as boolean

sub main()

    if asAny( nothing ) then
       debug.print("asAny( nothing )      returned true" )
    else
       debug.print("asAny( nothing )      returned false")
    end if

  ' --------------------------------------------------------

  ' if asAny( 0 ) then --> Compilie error: Type mismatch

  ' --------------------------------------------------------

    if asAny( vbNullString ) then
       debug.print("asAny( vbNullString ) returned true" )
    else
       debug.print("asAny( vbNullString ) returned false")
    end if

  ' --------------------------------------------------------

    dim text as string
    text = "Hello World"

    if asAny( text ) then
       debug.print("asAny( text )         returned true" )
    else
       debug.print("asAny( text )         returned false")
    end if

end sub
