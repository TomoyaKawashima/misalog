Option Explicit

Dim StartDate, EndDate, Res

If WScript.Arguments.Count <> 2 then
    WScript.echo("システム管理者に連絡してください.")
    WScript.Quit(-1)
End If
StartDate = WScript.Arguments(0)
EndDate = WScript.Arguments(1)
Res = CheckFormat(StartDate, EndDate)
WScript.Echo(Res)

Function CheckFormat(ByVal SDate1, ByVal EDate1)
    If Len(SDate1) <> CnLen(SDate1) Or Len(EDate1) <> CnLen(EDate1) Then
        MsgBox "Incorrect format."
        WScript.Quit(-1)
    End If
    CheckFormat = 1
End Function

Function CnLen(ByVal StrVal)
    Dim i, StrChr
    CnLen = 0
    If Trim(strVal) <> "" Then
        For i = 1 To Len(StrVal)
            StrChr = Mid(StrVal, i, 1)
            If (Asc(StrChr) And &HFF00) <> 0 Then
                CnLen = CnLen + 2
            Else
                CnLen = CnLen + 1
            End If
        Next
    End If
End Function