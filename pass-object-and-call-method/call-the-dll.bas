option explicit

declare sub callMethodOnObj                                                  _
        lib "c:\github\VBA-calls-DLL\pass-object-and-call-method\the.dll" (  _
           obj as object                                                     _
        )

sub main() ' {

    dim obj as new tq84

    callMethodOnObj obj

end sub ' }
