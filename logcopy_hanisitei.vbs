Option Explicit

Dim startdate: startdate = inputbox("ログを採取する期間を入力してください")
Dim enddate: enddate = inputbox("ログを採取する期間を入力してください")

Dim objfilesys: objfilesys = CreateObject("Scripting.FileSystemObject")
Dim objfolder
Dim objfile 
Dim objfiledate
Dim objfs: objfs = CreateObject("Scripting.FileSystemObject")

Dim copyfrom: copyfrom = "C:\test\aaa"
Dim copyfromfile
Dim copyto: copyto = "C:\test\bbb"
Dim copytofile

set objfolder = objfilesys.GetFolder(copyfrom)

for each objfile in objfolder.files 
    objfiledate = mid(objfile.name, 14, 8)
    if objfiledate >= startdate And objfiledate <= enddate Then
        call objfs.copyfile(copyfrom&objfile.name, copyto&objfile.name, true)
    End if
next