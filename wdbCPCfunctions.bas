Option Compare Database
Option Explicit

Function openCPCprojectFolder(projectNumber As String)
On Error GoTo err_handler

Dim mainFolder As String
mainFolder = "\\Srv-corp-nas01\quality\1 CPC\" & getYear(projectNumber) & " CPC Project Folder\"
If FolderExists(mainFolder) = False Then MkDir (mainFolder)

Dim FolName As String, FilePath As String, projectFolder As String
projectFolder = mainFolder & projectNumber
FolName = Dir(projectFolder & "*", vbDirectory)
FilePath = mainFolder & FolName

If Len(FolName) = 0 Then
    MkDir (projectFolder)
    FilePath = projectFolder
End If

openPath (FilePath)

Exit Function
err_handler:
    Call handleError("wdbCPCfunctions", "openCPCprojectFolder", Err.DESCRIPTION, Err.number)
End Function

Public Sub registerCPCUpdates(table As String, ID As Variant, column As String, oldVal As Variant, newVal As Variant, projectId As Long, Optional tag0 As String, Optional tag1 As String)
On Error GoTo err_handler

If (VarType(oldVal) = vbDate) Then oldVal = Format(oldVal, "mm/dd/yyyy")
If (VarType(newVal) = vbDate) Then newVal = Format(newVal, "mm/dd/yyyy")

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("tblCPC_UpdateTracking")

If Len(oldVal) > 255 Then oldVal = Left(oldVal, 255)
If Len(newVal) > 255 Then newVal = Left(newVal, 255)
If Len(tag0) > 100 Then newVal = Left(tag0, 100)
If Len(tag1) > 100 Then newVal = Left(tag1, 100)
If ID = "" Then ID = Null

With rs1
    .addNew
        !tableName = table
        !tableRecordId = ID
        !updatedBy = Environ("username")
        !updatedDate = Now()
        !columnName = column
        !previousData = StrQuoteReplace(CStr(Nz(oldVal, "")))
        !newData = StrQuoteReplace(CStr(Nz(newVal, "")))
        !projectId = projectId
        !dataTag0 = StrQuoteReplace(tag0)
        !dataTag1 = StrQuoteReplace(tag1)
    .Update
    .Bookmark = .lastModified
End With

rs1.Close
Set rs1 = Nothing
Set db = Nothing


Exit Sub
err_handler:
    Call handleError("wdbCPCfunctions", "registerCPCUpdates", Err.DESCRIPTION, Err.number)
End Sub

Function getYear(projectNumber As String)
On Error GoTo err_handler

    If Len(projectNumber) = 7 Then
        getYear = Left(Year(Now), 2) & Mid(projectNumber, 2, 2)
    Else
        getYear = Left(Year(Now), 2) & Mid(projectNumber, 3, 2)
    End If
    
Exit Function
err_handler:
    Call handleError("wdbCPCfunctions", "getYear", Err.DESCRIPTION, Err.number)
End Function