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
    Call handleError("wdbCPCfunctions", "openCPCprojectFolder", Err.Description, Err.number)
End Function