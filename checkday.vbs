Option Explicit

Sub main()
    Dim StartDate, EndDate
    Dim Res: Res = 0
    If WScript.Arguments.Count <> 2 then
        MsgBox "Please set two days."
        WScript.Quit(-1)
    End If
    StartDate = WScript.Arguments(0)
    EndDate = WScript.Arguments(1)
    Res = IsValidInput(StartDate, EndDate)
    Res = IsCorrectFormat(StartDate, EndDate)
    Res = IsCorrectDate(StartDate)
    Res = IsCorrectDate(EndDate)
    Res = IsValidPeriod(StartDate, EndDate)
    Res = IsLimit(StartDate, EndDate)
    WScript.Echo(ValidPeriodRes)
End Sub

' 始めと終わりの日付が数字で入力されているか
Function IsValidInput(Byval SDate1, EDate1)
    If IsNumeric(StartDate) = False Or IsNumeric(EndDate) = False Then
        MsgBox "This input isn't Numeric."
    End If
End Function

' 入力フォーマットが正しいか
Function IsCorrectFormat(ByVal SDate2, ByVal EDate2)
    IsCorrectFormat = True
    Dim InputLen: InputLen = 8 
    If Len(SDate2) <> ByteLen(SDate2) Or Len(EDate2) <> ByteLen(EDate2) Then
        MsgBox "Incorrect format."
        IsCorrectFormat = False
    End If
    If Len(SDate2) <> InputLen Or Len(EDate2) <> InputLen Then
        MsgBox "Incorrect format."
        IsCorrectFormat = False
    End If
End Function

Function ByteLen(ByVal StrVal)
    Dim i, StrChr
    ByteLen = 0
    If Trim(StrVal) <> "" Then
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

' 日付が存在するか
Function IsCorrectDate(ByVal InputDate)
    Dim Re, Mc
    IsCorrectDate = True
    set Re = CreateObject("VBScript.RegExp")
    Re.Pattern = "^(?!([02468][1235679]|[13579][01345789])000229)(([0-9]{4}(01|03|05|07|08|10|12)(0[1-9]|[12][0-9]|3[01]))|([0-9]{4}(04|06|09|11)(0[1-9]|[12][0-9]|30))|([0-9]{4}02(0[1-9]|1[0-9]|2[0-8]))|([0-9]{2}([02468][048]|[13579][26])0229))$"
    set Mc = Re.Execute(InputDate)
    If Mc.Count = 0 Then 
        IsCorrectDate = False
    End If
End Function

' 終わりの日付が始めの日付より後ろの日付になっているか
Function IsValidPeriod(Byval SDate3, Byval EDate3)
    IsValidPeriod = True
    If (EDate3 - SDate3) < 0 Then
        MsgBox "Invalid period."
        IsValidPeriod = False 
    End If 
End Function

' 取り出せる期間に収まっているか
Function IsLimit(ByVal SDate4, ByVal EDate4)
    IsLimit = True
    Dim LimitMonth: LimitMonth = 3
    Dim CheckDate, LimitDate, NewDate
    CheckDate = Mid(SDate4, 1, 4) & "/" & Mid(SDate4, 5, 2) & "/" & Mid(SDate4, 7, 2)
    LimitDate = DateAdd("m", LimitMonth, CheckDate)
    LimitDate = Replace(LimitDate, "/", "")
    If(LimitDate - EDate4 < 0) Then
        MsgBox "Out of bound."
        IsLimit = False
    End If 
End Function
    
    


