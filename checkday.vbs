Option Explicit

'カウンター
Class CallCount
    private Sub Class_Initialize()
    End Sub 
    private Sub Class_Terminate()
    End Sub
    Private Cnt
    Public Function Counter()
        Cnt = Cnt + 1
        Counter = Cnt   
    End Function
End Class

' エラーメッセージ用ラッパー
Class Result
    private Sub Class_Initialize()
    End Sub 
    private Sub Class_Terminate()
        WScript.Quit(0)
    End Sub
    Private Mes
    Public Property Let Message(ErrMes)
        Mes = ErrMes 
        WScript.Echo(Mes)
    End Property
End Class

main()

Sub main()
    Dim StartDate, EndDate, Res
    Dim LimitMonth: LimitMonth = 3
    Dim CallCount: CallCount = 0
    Dim ResObj
    Set ResObj = New CallCount
    If WScript.Arguments.Count <> 2 then
        ClsObj.Message = "Please set two days."
    End If

    If IsNumeric(LimitMonth) = False Then
        ClsObj.Message = "Please set LimitMonth."

    StartDate = WScript.Arguments(0)
    EndDate = WScript.Arguments(1)

' ここをなんとかしたい
    Res = IsValidInput(StartDate, EndDate)
    Call AfterProc(Res, ResObj.Counter)
    Res = IsCorrectFormat(StartDate, EndDate)
    Call AfterProc(Res, ResObj.Counter)
    Res = IsCorrectDate(StartDate, EndDate)
    Call AfterProc(Res, ResObj.Counter)
    Res = IsValidPeriod(StartDate, EndDate)
    Call AfterProc(Res, ResObj.Counter)
    Res = IsLimit(StartDate, EndDate, LimitMonth)
    Call AfterProc(Res, ResObj.Counter)
    WScript.Echo("success")
    Wscript.Quit(-1)
End Sub

' 始めと終わりの日付が数字で入力されているか
Function IsValidInput(Byval SDate1, EDate1)
    IsValidInput = True
    If IsNumeric(SDate1) = False Or IsNumeric(EDate1) = False Then
       IsValidInput = False
    End If
End Function

' 入力フォーマットが正しいか
Function IsCorrectFormat(ByVal SDate2, ByVal EDate2)
    IsCorrectFormat = True
    Dim InputLen: InputLen = 8 
    If Len(SDate2) <> ByteLen(SDate2) Or Len(EDate2) <> ByteLen(EDate2) Then
        IsCorrectFormat = False
    End If
    If Len(SDate2) <> InputLen Or Len(EDate2) <> InputLen Then
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
Function IsCorrectDate(ByVal SDate3, EDate3)
    Dim RegexSDate, RegexEDate, MatchCountSDate, MatchCountEDate
    IsCorrectDate = True
    Set RegexSDate = CreateObject("VBScript.RegExp")
    Set RegexEDate = CreateObject("VBScript.RegExp")
    RegexSDate.Pattern = "^(?!([02468][1235679]|[13579][01345789])000229)(([0-9]{4}(01|03|05|07|08|10|12)(0[1-9]|[12][0-9]|3[01]))|([0-9]{4}(04|06|09|11)(0[1-9]|[12][0-9]|30))|([0-9]{4}02(0[1-9]|1[0-9]|2[0-8]))|([0-9]{2}([02468][048]|[13579][26])0229))$"
    RegexEDate.Pattern = "^(?!([02468][1235679]|[13579][01345789])000229)(([0-9]{4}(01|03|05|07|08|10|12)(0[1-9]|[12][0-9]|3[01]))|([0-9]{4}(04|06|09|11)(0[1-9]|[12][0-9]|30))|([0-9]{4}02(0[1-9]|1[0-9]|2[0-8]))|([0-9]{2}([02468][048]|[13579][26])0229))$"
    Set MatchCountSDate = RegexSDate.Execute(SDate3)
    Set MatchCountEDate = RegexEDate.Execute(EDate3)
    If MatchCountSDate.Count = 0 Or MatchCountEDate.Count = 0 Then 
        IsCorrectDate = False
    End If
End Function

' 終わりの日付が始めの日付より後ろの日付になっているか
Function IsValidPeriod(Byval SDate4, Byval EDate4)
    IsValidPeriod = True
    If (EDate4 - SDate4) < 0 Then
        IsValidPeriod = False 
    End If 
End Function

' 取り出せる期間に収まっているか
Function IsLimit(ByVal SDate5, ByVal EDate5, ByVal LMonth)
    IsLimit = True
    Dim CheckDate, LimitDate, NewDate
    CheckDate = Mid(SDate5, 1, 4) & "/" & Mid(SDate5, 5, 2) & "/" & Mid(SDate5, 7, 2)
    LimitDate = DateAdd("m", LMonth, CheckDate)
    LimitDate = Replace(LimitDate, "/", "")
    If(LimitDate - EDate5 < 0) Then
        IsLimit = False
    End If 
End Function

'関数呼び出し後処理
Sub AfterProc(ByVal Res, ByVal CallCount)
    If Res <> -1 Then
        ErrMsg(CallCount)
    End If
End Sub

' エラーメッセージ    
Function ErrMsg(ByVal CallCount)
Dim ClsObj
Set ClsObj = New Result
    Select Case CallCount
        Case 1
            ClsObj.Message = "This input isn't numeric."
        Case 2 
            ClsObj.Message = "Incorrect format."
        Case 3
            ClsObj.Message = "This Date isn't found."
        Case 4 
            ClsObj.Message = "Invalid period."
        Case 5 
            ClsObj.Message = "Out of bound."
    End Select
End Function


