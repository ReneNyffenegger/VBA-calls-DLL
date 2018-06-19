option explicit

declare sub swap_double_ptr                        _
        lib "c:\github\VBA-calls-DLL\c\the.dll" (  _
           byRef a as double,  _
           byRef b as double   _
        )

declare function diff_double                       _
        lib "c:\github\VBA-calls-DLL\c\the.dll" (  _
           byVal a as double,  _
           byVal b as double   _
        ) as double

declare sub to_upper_char_ptr                      _
        lib "c:\github\VBA-calls-DLL\c\the.dll" (  _
           byVal p as string  _
        )

sub main()

  dim a as double
  dim b as double
  dim d as double

  a = 123.45
  b =  12.34

  d = diff_double(a, b)

  msgBox "d = " & d

  call swap_double_ptr(a, b)

  msgBox "After swap: a = " & a & ", b = " & b

  dim str as string
  str = "Foo Bar Baz"
  to_upper_char_ptr str
  msgBox str

end sub
