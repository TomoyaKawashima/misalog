Option Explicit

Dim StartDate, EndDate

if WScript.Arguments.Count <> 2 then
    WScript.echo("システム管理者に連絡してください.")
    WScript.Quit(-1)
end if

StartDate = WScript.Arguments(0)
EndDate = WScript.Arguments(1)

Wscript.Echo(EndDate - StartDate)