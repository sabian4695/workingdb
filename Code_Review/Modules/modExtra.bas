Option Compare Database

Function disableShift()

Dim db, acc
Set acc = CreateObject("Access.Application")
'Set db = acc.DBEngine.OpenDatabase("\\data\mdbdata\WorkingDB\build\Commands\Misc_Commands\WorkingDB_SummaryEmail.accdb", False, False)
'Set db = acc.DBEngine.OpenDatabase("H:\dev\WorkingDB_SummaryEmail.accdb", False, False)
Set db = acc.DBEngine.OpenDatabase("C:\workingdb\WorkingDB_ghost.accdb", False, False)


db.Properties("AllowByPassKey") = True

db.Close
Set db = Nothing

End Function

Function disableShift_FE()

Dim db, acc, fso
Set acc = CreateObject("Access.Application")

Set fso = CreateObject("Scripting.FileSystemObject")

Dim repoLoc As String
repoLoc = fso.GetParentFolderName(CurrentProject.Path) & "\Front_End\WorkingDB_FE.accdb"

Set db = acc.DBEngine.OpenDatabase(repoLoc, False, False)


db.Properties("AllowByPassKey") = True

db.Close
Set db = Nothing

End Function