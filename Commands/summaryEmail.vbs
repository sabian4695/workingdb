Dim oApplication
Set oApplication = CreateObject("Access.Application")
oApplication.OpenCurrentDatabase "\\data\mdbdata\WorkingDB\build\workingdb_summaryemail\WorkingDB_summaryEmail.accde"
'set oApplication = Nothing