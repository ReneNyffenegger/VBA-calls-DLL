option explicit

declare sub       setCallbackFunction lib "c:\github\VBA-calls-DLL\callback\the.dll" (byVal addrCallback as long)
declare function callCallbackFunction lib "c:\github\VBA-calls-DLL\callback\the.dll" (byVal i            as long) as long

sub main() ' {

    setCallbackFunction(vba.int(addressOf callback))

    debug.print "Return value for 1 = " & callCallbackFunction(1)
    debug.print "Return value for 2 = " & callCallbackFunction(2)
    debug.print "Return value for 3 = " & callCallbackFunction(3)
    debug.print "Return value for 4 = " & callCallbackFunction(4)

end sub ' }

function callBack(byVal num as long, byVal txt as string) as long ' {

    debug.print "Callback was called, num = " & num & ", txt = " & txt

    callBack = num + 5

end function ' }
