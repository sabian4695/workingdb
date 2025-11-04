on error resume next
Dim oApplication
Set oApplication = CreateObject("Access.Application")
oApplication.OpenCurrentDatabase "C:\workingdb\WorkingDB_ghost.accde"
set oApplication = Nothing