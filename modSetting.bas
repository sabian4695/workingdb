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
    Dim i As Integer
    Set db = CurrentDb()
    Set rs1 = db.OpenRecordset("tblPLMsettings", dbOpenSnapshot)
    
For i = 1 To 10
    Select Case i
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
Next i

rs1.Close
Set rs1 = Nothing
    
fncRead = True
End Function

Public Function fncCheck() As Boolean
    fncCheck = False

    If Trim(gstrSendToPath) = "" Then
        Exit Function
    End If
    If Trim(gstrServerName) = "" Then
        Exit Function
    End If
    If Trim(gstrUserName) = "" Then
        Exit Function
    End If
    If Trim(gstrPassword) = "" Then
        Exit Function
    End If
    If Trim(gstrDBName) = "" Then
        Exit Function
    End If
    If Trim(gstr3dexCacheDir) = "" Then
        Exit Function
    End If
    If Trim(gstrInputCheck) = "" Then
        Exit Function
    End If
    If Trim(gstrSaveAsNewName) = "" Then
        Exit Function
    End If
    If Trim(gstrAutoInput) = "" Then
        Exit Function
    End If
    If Trim(gstrUnsetMaterialGrade) = "" Then
        Exit Function
    End If
    fncCheck = True
End Function

Public Function fncCheckSendToPath() As Boolean
    fncCheckSendToPath = False
    
    If fncCheckDir(gstrSendToPath) = True Then
        fncCheckSendToPath = True
    End If
End Function

Public Function fncCheck3dexCacheDir() As Boolean
    fncCheck3dexCacheDir = False
    
    If fncCheckDir(gstr3dexCacheDir) = True Then
        fncCheck3dexCacheDir = True
    End If
End Function

Private Function fncCheckDir(ByVal istrPath As String) As Boolean
    fncCheckDir = False
    
    Dim strResult As String
    strResult = ""
    
    On Error Resume Next
    strResult = Dir(istrPath, vbDirectory)
    On Error GoTo 0
    
    If strResult <> "" Then
        fncCheckDir = True
    End If
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