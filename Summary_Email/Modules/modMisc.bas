Option Compare Database
Option Explicit

Public Function StrQuoteReplace(strValue)
StrQuoteReplace = Replace(strValue, "'", "''")
End Function

Function emailContentGen(subject As String, Title As String, subTitle As String, primaryMessage As String, detail1 As String, detail2 As String, detail3 As String) As String
emailContentGen = subject & "," & Title & "," & subTitle & "," & primaryMessage & "," & detail1 & "," & detail2 & "," & detail3
End Function

Function getEmail(userName As String) As String
On Error GoTo tryOracle
Dim rsPermissions As Recordset
Set rsPermissions = CurrentDb().OpenRecordset("SELECT * from tblPermissions WHERE user = '" & userName & "'")
getEmail = rsPermissions!userEmail
rsPermissions.Close

Exit Function
tryOracle:
Dim rsEmployee As Recordset
Set rsEmployee = CurrentDb().OpenRecordset("SELECT FIRST_NAME, LAST_NAME, EMAIL_ADDRESS FROM APPS_XXCUS_USER_EMPLOYEES_V WHERE USER_NAME = '" & StrConv(userName, vbUpperCase) & "'")
getEmail = Nz(rsEmployee!EMAIL_ADDRESS, "")
End Function