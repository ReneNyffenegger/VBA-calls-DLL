option explicit

declare sub fourLongs                                    _
        lib "c:\github\VBA-calls-DLL\c\fourLongs.dll" (  _
               byVal a as longPtr                     ,  _
               byRef b as longPtr                     ,  _
               byVal c as any                         ,  _
               byRef d as any                         )

sub main() ' {

    dim varLong   as long
    dim varString as string

    varLong   = 42
    varString ="Hello world"

    debug.print "varPtr(varLong)   = " & varPtr(varLong  )
    debug.print "varPtr(varString) = " & varPtr(varString)
    debug.print "strPtr(varString) = " & strPtr(varString)

    fourLongs        varLong   ,        varLong   ,        varLong   ,        varLong
    fourLongs strPtr(varString), strPtr(varString), strPtr(varString), strPtr(varString)
    fourLongs varPtr(varString), varPtr(varString), varPtr(varString), varPtr(varString)
    fourLongs vbNull           , vbNull           , vbNull           , vbNull
    fourLongs vbNull           , vbNull           , vbNullString     , vbNullString

end sub ' }
