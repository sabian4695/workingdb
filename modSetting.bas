Option Explicit

Public gstrSendToPath As String
Public gstrServerName As String
Public gstrUserName As String
Public gstrPassword As String
Public gstrDBName As String
Public gstrOldServerName As String
Public gstrOldUserName As String
Public gstrOldPassword As String
Public gstrOldDBName As String
Public gstrExcelPassword As String
Public gstr3dexCacheDir As String
Public gstrInputCheck As String
Public gstrSaveAsNewName As String
Public gstrAutoInput As String
Public gstrUnsetMaterialGrade As String

Public Function fncRead() As Boolean
fncRead = False
    On Error Resume Next
    On Error GoTo 0
    
    Dim rs1 As Recordset
    Dim db As Database
    Dim I As Integer
    Set db = CurrentDb()
    Set rs1 = db.OpenRecordset("tblPLMsettings", dbOpenSnapshot)
    
For I = 1 To 10
    Select Case I
        Case 1
            gstrSendToPath = rs1("Value")
        Case 2
            gstrServerName = rs1("Value")
        Case 3
            gstrUserName = rs1("Value")
        Case 4
            gstrPassword = rs1("Value")
        Case 5
            gstrDBName = rs1("Value")
        Case 6
            gstr3dexCacheDir = rs1("Value")
        Case 7
            gstrInputCheck = rs1("Value")
        Case 8
            gstrSaveAsNewName = rs1("Value")
        Case 9
            gstrAutoInput = rs1("Value")
        Case 10
            gstrUnsetMaterialGrade = rs1("Value")
    End Select
rs1.MoveNext
Next I

rs1.Close
Set rs1 = Nothing
Set db = Nothing
    
fncRead = True
End Function

Public Function fncCheck() As Boolean
    fncCheck = False
    If Trim(gstrSendToPath) = "" Then Exit Function
    If Trim(gstrServerName) = "" Then Exit Function
    If Trim(gstrUserName) = "" Then Exit Function
    If Trim(gstrPassword) = "" Then Exit Function
    If Trim(gstrDBName) = "" Then Exit Function
    If Trim(gstr3dexCacheDir) = "" Then Exit Function
    If Trim(gstrInputCheck) = "" Then Exit Function
    If Trim(gstrSaveAsNewName) = "" Then Exit Function
    If Trim(gstrAutoInput) = "" Then Exit Function
    If Trim(gstrUnsetMaterialGrade) = "" Then Exit Function
    fncCheck = True
End Function

Public Sub Terminate()
    gstrSendToPath = ""
    gstrServerName = ""
    gstrUserName = ""
    gstrPassword = ""
    gstrDBName = ""
    gstrOldServerName = ""
    gstrOldUserName = ""
    gstrOldPassword = ""
    gstrOldDBName = ""
    gstrExcelPassword = ""
    gstr3dexCacheDir = ""
    gstrInputCheck = ""
    gstrSaveAsNewName = ""
    gstrAutoInput = ""
    gstrUnsetMaterialGrade = ""
End Sub