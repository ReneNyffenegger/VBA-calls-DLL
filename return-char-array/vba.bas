option explicit

declare ptrSafe function charArray  _
        lib "the.dll" (             _
        ) as longPtr

declare ptrSafe function wcharArray _
        lib "the.dll" (             _
        ) as longPtr

declare ptrSafe function bstr       _
        lib "the.dll" (             _
        ) as string

declare ptrSafe function bstr_c     _
        lib "the.dll" (             _
        ) as string


private const CP_UTF8 as long = 65001
declare ptrsafe function MultiByteToWideChar lib "kernel32" ( _
  byVal CodePage        as long   , _
  byVal dwFlags         as long   , _
  byVal lpMultiByteStr  as longPtr, _
  byVal cbMultiByte     as long   , _
  byVal lpWideCharStr   as longPtr, _
  byVal cchWideChar     as long     _
) as long

declare ptrsafe function lstrlenW lib "kernel32" ( _
  byVal lpSTring        as longPtr  _
) as long

declare ptrsafe function lstrcpyW lib "kernel32" ( _
  byVal lpString1       as longPtr,  _
  byVal lpString2       as longPtr   _
) as longPtr

'#if Win64 then
function utf8PtrToString(byVal pUtf8string as longPtr) as string ' {
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

end function ' }

function wcharPtrToString(byVal wcharPtr as longPtr) as string ' {
    dim cSize  as long

    csize = lstrlenW(wcharPtr)
    wcharPtrToString = string(cSize, "*")
    lstrcpyW strPtr(wcharPtrToString), wcharPtr

end function ' }

sub main() ' {

    dim  s as string
    dim ws as string

    s = utf8PtrToString(charArray())
    debug.print("s = " & s)

    ws = wcharPtrToString(wcharArray())
    debug.print("ws = " & ws)

    debug.print("bstr = " & bstr())

    debug.print("bstr_c = " & bstr_c())

end sub ' }
