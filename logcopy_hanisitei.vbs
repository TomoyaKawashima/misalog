Option explicit

Dim startdate
Dim enddate

startdate = inputbox("ログを採取する期間を入力してください")
enddate = inputbox("ログを採取する期間を入力してください")

Dim objfilesys
Dim objfolder
Dim objfile 
Dim objfiledate
Dim objfs 

Dim copyfrom
Dim copyfromfile
Dim copyto
Dim copytofile

copyfrom = "C:\test\aaa"
copyto = "C:\test\bbb"

set objfilesys = CreateObject("scripting.filesystemobject")
set objfolder = objfilesys.getfolder(copyfrom)
set objfs = createobject("scripting.filesystemobject")

for each objfile in objfolder.files 
    objfiledate = mid(objfile.name, 14, 8)
    if objfiledate >= startdate And objfiledate <= enddate Then
        call objfs.copyfile(copyfrom&objfile.name, copyto&objfile.name, true)
    End if
next