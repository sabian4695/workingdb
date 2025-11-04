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