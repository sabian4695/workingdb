Option Compare Database
Option Explicit

Function splitString(a, b, c) As String
    On Error GoTo errorCatch
    splitString = Split(a, b)(c)
    Exit Function
errorCatch:
    splitString = ""
End Function

Public Function grabGatePlannedDate(partNumber As String, gateNum As Long) As Date
On Error Resume Next

Dim db As Database
Dim rs As Recordset

Set db = CurrentDb()

Set rs = db.OpenRecordset("SELECT * FROM tblPartGates WHERE partNumber = '" & partNumber & "' AND gateTitle Like 'G" & gateNum & "*'")

If rs.RecordCount = 0 Then GoTo skip

grabGatePlannedDate = rs!plannedDate

skip:
On Error Resume Next
rs.Close
Set rs = Nothing

Set db = Nothing

End Function