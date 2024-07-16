Option Compare Database
Option Explicit

Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal lpnShowCmd As Long) As Long

Public Sub openPath(Path)
On Error GoTo err_handler

CreateObject("Shell.Application").open CVar(Path)

Exit Sub
err_handler:
    Call handleError("modGlobal", "openPath", Err.Description, Err.number)
End Sub

Function replaceDriveLetters(linkInput) As String

replaceDriveLetters = Replace(linkInput, "N:\", "\\ncm-fs2\data\Department\")
replaceDriveLetters = Replace(linkInput, "T:\", "\\design\data\")
replaceDriveLetters = Replace(linkInput, "S:\", "\\nas01\allshare\")

End Function

Function addLastSlash(linkString As String) As String
On Error GoTo err_handler
addLastSlash = linkString
If Right(addLastSlash, 1) <> "\" Then addLastSlash = addLastSlash & "\"

Exit Function
err_handler:
    Call handleError("modGlobal", "addLastSlash", Err.Description, Err.number)
End Function

Function createShortcut(lnkLocation As String, targetLocation As String, shortcutName As String)
On Error GoTo err_handler

With CreateObject("WScript.Shell").createShortcut(lnkLocation & ".lnk")
    .TargetPath = targetLocation
    .Description = shortcutName
    .save
End With

Exit Function
err_handler:
    Call handleError("modGlobal", "createShortcut", Err.Description, Err.number)
End Function

Public Sub checkMkDir(mainFolder, partNum, Optional variableVal)
On Error GoTo err_handler
Dim FolName As String, fullPath As String

If variableVal = "*" Then
    FolName = Dir(mainFolder & partNum & "*", vbDirectory)
Else
    FolName = partNum
End If

If FolName = "" Then FolName = partNum

fullPath = mainFolder & FolName

If Len(partNum) = 5 Or (partNum Like "D*" And Len(partNum) = 6) Then
    If FolderExists(fullPath) Then
        Call openPath(fullPath)
        Exit Sub
    End If
    If MsgBox("This folder does not exist. Create folder?", vbYesNo, "Folder Does Not Exist") = vbYes Then
        MkDir (fullPath)
        MsgBox "Folder Created. Going to New Folder.", vbOKOnly, "Folder Created"
        Call openPath(fullPath)
    Else
        If MsgBox("Folder Not Created. Do you want to go to the main folder?", vbYesNo, "Folder Not Created") = vbYes Then Call openPath(mainFolder)
        Exit Sub
    End If
Else
    Call openPath(mainFolder)
End If
Exit Sub
err_handler:
    Call handleError("basGlobal", "checkMkDir", Err.Description, Err.number)
End Sub

Function mainFolder(sName As String) As String
mainFolder = DLookup("[Link]", "tblLinks", "[btnName] = '" & sName & "'")
End Function

Function FolderExists(sFile As Variant) As Boolean
FolderExists = False
If IsNull(sFile) Then Exit Function
If Dir(sFile, vbDirectory) <> "" Then FolderExists = True
End Function

Public Function zeros(partNum, Amount As Variant)
    If (Amount = 2) Then
        zeros = Left(partNum, 3) & "00\"
    ElseIf (Amount = 3) Then
        zeros = Left(partNum, 2) & "000\"
    End If
End Function

Function openDocumentHistoryFolder(partNum)
Dim thousZeros, hundZeros
Dim mainPath, FolName, strFilePath, prtFilePath, dPath As String

If partNum Like "D*" Then
    Call checkMkDir(mainFolder("DocHisD"), partNum, "*")
ElseIf partNum Like "[A-Z][A-Z]##[A-Z]##[A-Z]" Or partNum Like "[A-Z][A-Z]##[A-Z]##" Or partNum Like "##[A-Z]##" Then
    'Examples: AB11A76A or AB11A76 or 11A76
    If Not partNum Like "##[A-Z]##" Then
        partNum = Mid(partNum, 3, 5)
    End If
    mainPath = mainFolder("ncmDrawingMaster")
    prtFilePath = mainPath & Left(partNum, 3) & "00\" & partNum & "\"
    strFilePath = prtFilePath & "Documents"
    
    If FolderExists(strFilePath) = True Then
        Call openPath(strFilePath)
    Else
        DoCmd.OpenForm "frmCreateDesignFolders"
    End If
Else
    thousZeros = Left(partNum, 2) & "000\"
    hundZeros = Left(partNum, 3) & "00\"
    mainPath = mainFolder("docHisSearch")
    prtFilePath = mainPath & thousZeros & hundZeros
    FolName = Dir(prtFilePath & partNum & "*", vbDirectory)
    strFilePath = prtFilePath & FolName
    
    If Len(partNum) = 5 Or Right(partNum, 1) = "P" Then
        If Len(FolName) = 0 Then
            DoCmd.OpenForm "frmCreateDesignFolders"
        Else
            Call openPath(strFilePath)
        End If
    Else
        Call openPath(mainPath)
    End If
End If
End Function

Function openModelV5Folder(partNumOriginal)
Dim partNum, thousZeros, hundZeros, FolName, mainfolderpath, strFilePath, prtpath, dPath

partNum = partNumOriginal & "_"
If partNum Like "D*" Then
    Call checkMkDir(mainFolder("ModelV5D"), Left(partNum, Len(partNum) - 1), "*")
ElseIf Left(partNum, 8) Like "[A-Z][A-Z]##[A-Z]##[A-Z]" Or Left(partNum, 7) Like "[A-Z][A-Z]##[A-Z]##" Or Left(partNum, 5) Like "##[A-Z]##" Then
    'Examples: AB11A76A or AB11A76 or 11A76
    partNum = partNumOriginal
    If Not partNum Like "##[A-Z]##" Then
        partNum = Mid(partNum, 3, 5)
    End If
    mainfolderpath = mainFolder("ncmDrawingMaster")
    prtpath = mainfolderpath & Left(partNum, 3) & "00\" & partNum & "\"
    strFilePath = prtpath & "CATIA"
    
    If FolderExists(strFilePath) = True Then
        Call openPath(strFilePath)
    Else
        DoCmd.OpenForm "frmCreateDesignFolders"
    End If
Else
    thousZeros = Left(partNum, 2) & "000\"
    hundZeros = Left(partNum, 3) & "00\"
    mainfolderpath = mainFolder("modelV5search")
    prtpath = mainfolderpath & thousZeros & hundZeros
tryagain:
    FolName = Dir(prtpath & partNum & "*", vbDirectory)
    strFilePath = prtpath & FolName
    
    If Len(partNumOriginal) = 5 Or partNumOriginal Like "*P" Then
        If Len(FolName) = 0 Then
            If partNum Like "*_" Then
                partNum = Left(partNum, 5)
                GoTo tryagain
            End If
            DoCmd.OpenForm "frmCreateDesignFolders"
        Else
            Call openPath(strFilePath)
        End If
    Else
        Call openPath(mainfolderpath)
    End If
End If
End Function