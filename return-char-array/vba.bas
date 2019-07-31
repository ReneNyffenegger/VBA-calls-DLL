option explicit

declare ptrSafe function charArray _
        lib "the.dll" (            _
        ) as longPtr

private const CP_UTF8 as long = 65001
declare ptrsafe function MultiByteToWideChar Lib "kernel32" ( _
  byVal CodePage        as long   , _
  byVal dwFlags         as long   , _
  byVal lpMultiByteStr  as longPtr, _
  byVal cbMultiByte     as long   , _
  byVal lpWideCharStr   as longPtr, _
  byVal cchWideChar     as long     _
) as long

'#if Win64 Then
function utf8PtrToString(byVal pUtf8string as longPtr) as string
' #Else
'   function utf8PtrToString(byVal pUtf8string as long   ) as string
' #End if
'
'   Found @ https://github.com/govert/SQLiteForExcel/blob/master/Source/SQLite3VBAModules/Sqlite3_64.bas
'
    dim buf    as string
    dim cSize  as long
    dim mbVal  as long
    
    cSize = MultiByteToWideChar(CP_UTF8, 0, pUtf8String, -1, 0, 0)
  ' cSize includes the terminating null character
    if cSize <= 1 Then
       utf8PtrToString = ""
       exit function
    end if
    
    utf8PtrToString = string(cSize - 1, "*") ' and a termintating null char.
    mbVal = MultiByteToWideChar(CP_UTF8, 0, pUtf8String, -1, strPtr(utf8PtrToString), cSize)

    if mbVal = 0 then
       err.raise 1000, "Error", "MultiByteToWideChar failed"
    end if

end function


sub main()

    dim s as string

    s = utf8PtrToString(charArray())

    debug.print("s = " & s)

end sub
