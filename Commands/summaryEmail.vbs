Dim oApplication
Set oApplication = CreateObject("Access.Application")
oApplication.OpenCurrentDatabase "\\data\mdbdata\WorkingDB\prod-FE\prod-functions\WorkingDB_SummaryEmail.accde"
'set oApplication = Nothing