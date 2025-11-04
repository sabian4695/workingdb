Option Explicit

Const acCmdCloseAll = &H286
Const acCmdCompileAndSaveAllModules = &H7E

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

reserveDev

Function reserveDev()
    Dim devLoc, strUser
    strUser = CreateObject("WScript.Network").UserName
    devLoc = "H:\dev\WorkingDB_" & strUser & "_dev.accdb"
    fso.CopyFile "\\data\mdbdata\WorkingDB\prod-FE\WorkingDB_FE.accdb", devLoc 

	Dim db, acc
	set acc = CreateObject("Access.Application")
	set db = acc.DBEngine.OpenDatabase(devLoc, False, False)

	db.Properties("AllowByPassKey") = True

End Function