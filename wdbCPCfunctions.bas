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

Public Sub registerCPCUpdates(table As String, ID As Variant, column As String, oldVal As Variant, newVal As Variant, Optional tag0 As String, Optional tag1 As String)
On Error GoTo err_handler

Dim sqlColumns As String, sqlValues As String

If (VarType(oldVal) = vbDate) Then
    oldVal = Format(oldVal, "mm/dd/yyyy")
End If

If (VarType(newVal) = vbDate) Then
    newVal = Format(newVal, "mm/dd/yyyy")
End If

If (IsNull(oldVal)) Then
    oldVal = ""
End If

If (IsNull(newVal)) Then
    newVal = ""
End If

sqlColumns = "(tableName,tableRecordId,updatedBy,updatedDate,columnName,previousData,newData,dataTag0"
                    
sqlValues = " values ('" & table & "', '" & ID & "', '" & Environ("username") & "', '" & Now() & "', '" & column & "', '" & StrQuoteReplace(CStr(oldVal)) & "', '" & StrQuoteReplace(CStr(newVal)) & "','" & tag0 & "'"

If (IsNull(tag1)) Then
    sqlColumns = sqlColumns & ")"
    sqlValues = sqlValues & ");"
Else
    sqlColumns = sqlColumns & ",dataTag1)"
    sqlValues = sqlValues & ",'" & tag1 & "');"
End If

dbExecute "INSERT INTO tblCPC_UpdateTracking" & sqlColumns & sqlValues

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