Option Explicit

Dim StartDate, EndDate, CorrectFormatRes, CorrectDateRes

If WScript.Arguments.Count <> 2 then
    MsgBox "Please set two days."
    WScript.Quit(-1)
End If
StartDate = WScript.Arguments(0)
EndDate = WScript.Arguments(1)
CorrectFormatRes = IsCorrectFormat(StartDate, EndDate)
CorrectDateRes = IsCorrectDate(EndDate)
WScript.Echo(CorrectDateRes)

Function IsCorrectFormat(ByVal SDate1, ByVal EDate1)
    IsCorrectFormat = 0
    If Len(SDate1) <> ByteLen(SDate1) Or Len(EDate1) <> ByteLen(EDate1) Then
        MsgBox "Incorrect format."
        ' WScript.Quit(-1)
        IsCorrectFormat = -1
    End If
    If Len(SDate1) <> 8 Or Len(EDate1) <> 8 Then
        MsgBox "Incorrect format."
        ' WScript.Quit(-1)
        IsCorrectFormat = -1
    End If
End Function

Function ByteLen(ByVal StrVal)
    Dim i, StrChr
    ByteLen = 0
    If Trim(strVal) <> "" Then
        For i = 1 To Len(StrVal)
            StrChr = Mid(StrVal, i, 1)
            If (Asc(StrChr) And &HFF00) <> 0 Then
                ByteLen = ByteLen + 2
            Else
                ByteLen = ByteLen + 1
            End If
        Next
    End If
End Function

Function IsCorrectDate(ByVal InputDate)
    Dim Re, Mc
    IsCorrectDate = 0
    set Re = CreateObject("VBScript.RegExp")
    Re.Pattern = "^(?!([02468][1235679]|[13579][01345789])000229)(([0-9]{4}(01|03|05|07|08|10|12)(0[1-9]|[12][0-9]|3[01]))|([0-9]{4}(04|06|09|11)(0[1-9]|[12][0-9]|30))|([0-9]{4}02(0[1-9]|1[0-9]|2[0-8]))|([0-9]{2}([02468][048]|[13579][26])0229))$"
    set Mc = Re.Execute(InputDate)
    If Mc.Count = 0 Then 
        IsCorrectDate = -1
    End If
End Function