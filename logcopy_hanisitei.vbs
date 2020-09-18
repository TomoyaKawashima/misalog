Option Explicit

Dim startdate: startdate = inputbox("ログを採取する期間を入力してください")
Dim enddate: enddate = inputbox("ログを採取する期間を入力してください")

Dim WshShell
Dim Flg: Flg = 0
Set WshShell = WScript.CreateObject("WScript.Shell")

Do While Flg <> -1
    If startdate = "" And enddate = "" Then
        WScript.Quit(0)
    End If 
    WScript.Echo(startdate)
    WScript.Echo(enddate)
    Flg =  WshShell.Run("checkday.vbs" & " " & startdate & " " & enddate, 0, True)
    If Flg <> -1 Then 
        startdate = inputbox("ログを採取する期間を入力してください")
        enddate = inputbox("ログを採取する期間を入力してください")
    End If
Loop

Dim objfilesys
Set objfilesys = CreateObject("Scripting.FileSystemObject")
Dim objfolder
Dim objfile 
Dim objfiledate
Dim objfs
Set objfs = CreateObject("Scripting.FileSystemObject")

Dim copyfrom: copyfrom = "D:\ic21\NS01\var\010.004.010.161\journal\"
Dim copyfromfile
Dim copyto: copyto = "D:\mw\Shunsaku\journal\"
Dim copytofile

set objfolder = objfilesys.GetFolder(copyfrom)

' 計算量がO(n)、とりあえず取り出して入ってるか調べる
for each objfile in objfolder.files 
    objfiledate = mid(objfile.name, 14, 8)
    if objfiledate >= startdate And objfiledate <= enddate Then
        call objfs.copyfile(copyfrom&objfile.name, copyto&objfile.name, true)
    End if
next
