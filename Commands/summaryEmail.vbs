Dim oApplication
Set oApplication = CreateObject("Access.Application")
oApplication.OpenCurrentDatabase "\\data\mdbdata\WorkingDB\build\Repo\Summary_Email\WorkingDB_summaryEmail.accde"
'set oApplication = Nothing