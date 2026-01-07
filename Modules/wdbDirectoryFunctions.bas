Option Compare Database
Option Explicit

Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal lpnShowCmd As Long) As Long

Public Sub openPath(Path)
On Error GoTo Err_Handler

CreateObject("Shell.Application").open CVar(Path)

Exit Sub
Err_Handler:
    Call handleError("wdbDirectoryFunctions", "openPath", Err.DESCRIPTION, Err.number)
End Sub

Function replaceDriveLetters(linkInput) As String
On Error GoTo Err_Handler

replaceDriveLetters = Replace(linkInput, "N:\", "\\ncm-fs2\data\Department\")
replaceDriveLetters = Replace(linkInput, "T:\", "\\design\data\")
replaceDriveLetters = Replace(linkInput, "S:\", "\\nas01\allshare\")

Exit Function
Err_Handler:
    Call handleError("wdbDirectoryFunctions", "replaceDriveLetters", Err.DESCRIPTION, Err.number)
End Function

Function addLastSlash(linkString As String) As String
On Error GoTo Err_Handler

addLastSlash = linkString
If Right(addLastSlash, 1) <> "\" Then addLastSlash = addLastSlash & "\"

Exit Function
Err_Handler:
    Call handleError("wdbDirectoryFunctions", "addLastSlash", Err.DESCRIPTION, Err.number)
End Function

Function createShortcut(lnkLocation As String, targetLocation As String, shortcutName As String)
On Error GoTo Err_Handler

If shortcutName <> "" Then shortcutName = " - " & shortcutName

With CreateObject("WScript.Shell").createShortcut(lnkLocation & shortcutName & ".lnk")
    .TargetPath = targetLocation
    .DESCRIPTION = shortcutName
    .save
End With

Exit Function
Err_Handler:
    Call handleError("wdbDirectoryFunctions", "createShortcut", Err.DESCRIPTION, Err.number)
End Function

Sub ListFilesInFolderAndSubfolders(folderPath As String)
    Dim fso As Object
    Dim startFolder As Object

    TempVars.Add "tStamp", Timer

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set startFolder = fso.GetFolder(folderPath)

    ProcessFolder startFolder
    
    checkTime ("DONE")
End Sub

Sub ProcessFolder(ByVal currentFolder As Object)
    Dim subFolder As Object
    Dim file As Object
    
    
    ' Loop through each file in the current folder
    For Each file In currentFolder.Files
        'Debug.Print file.Path ' Or perform other actions with the file (e.g., add to a list)
        dbExecute ("INSERT INTO tblSessionVariables(searchHistory) VALUES('" & file.Path & "')")
    Next file

    For Each subFolder In currentFolder.SubFolders
        ProcessFolder subFolder
    Next subFolder
End Sub

Public Function checkMkDir(mainFolder, partNum, Optional variableVal = "", Optional openBool As Boolean = True) As String
On Error GoTo Err_Handler
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
        checkMkDir = fullPath
    Else
        If MsgBox("This folder does not exist. Create folder?", vbYesNo, "Folder Does Not Exist") = vbNo Then
            If MsgBox("Folder Not Created. Do you want to go to the main folder?", vbYesNo, "Folder Not Created") = vbYes Then checkMkDir = mainFolder
        Else
            MkDir (fullPath)
            checkMkDir = fullPath
        End If
    End If
Else
    checkMkDir = mainFolder
End If

If openBool Then openPath (checkMkDir)

Exit Function
Err_Handler:
    Call handleError("wdbDirectoryFunctions", "checkMkDir", Err.DESCRIPTION, Err.number)
End Function

Function mainFolder(sName As String) As String
On Error GoTo Err_Handler

mainFolder = DLookup("[Link]", "tblLinks", "[btnName] = '" & sName & "'")

Exit Function
Err_Handler:
    Call handleError("wdbDirectoryFunctions", "mainFolder", Err.DESCRIPTION, Err.number)
End Function

Function FolderExists(sFile As Variant) As Boolean
On Error GoTo Err_Handler

FolderExists = False
If IsNull(sFile) Then Exit Function
If Dir(sFile, vbDirectory) <> "" Then FolderExists = True

Exit Function
Err_Handler:
    If Err.number = 52 Then Exit Function
    Call handleError("wdbDirectoryFunctions", "FolderExists", Err.DESCRIPTION, Err.number)
End Function

Public Function zeros(partNum, Amount As Variant)
On Error GoTo Err_Handler

    If (Amount = 2) Then
        zeros = Left(partNum, 3) & "00\"
    ElseIf (Amount = 3) Then
        zeros = Left(partNum, 2) & "000\"
    End If
    
Exit Function
Err_Handler:
    Call handleError("wdbDirectoryFunctions", "zeros", Err.DESCRIPTION, Err.number)
End Function

Function openDocumentHistoryFolder(partNum, Optional openBool As Boolean = True) As String
On Error GoTo Err_Handler
openDocumentHistoryFolder = ""

Dim thousZeros, hundZeros
Dim mainPath, FolName, strFilePath, prtFilePath, dPath As String
Dim exists As Boolean
exists = True

Select Case True
    Case partNum Like "D*"
        openDocumentHistoryFolder = checkMkDir(mainFolder("DocHisD"), partNum, "*", False)
    Case partNum Like "[A-Z][A-Z]##[A-Z]##[A-Z]" Or partNum Like "[A-Z][A-Z]##[A-Z]##" Or partNum Like "##[A-Z]##"
        'Examples: AB11A76A or AB11A76 or 11A76
        If Not partNum Like "##[A-Z]##" Then partNum = Mid(partNum, 3, 5)
        mainPath = mainFolder("ncmDrawingMaster")
        prtFilePath = mainPath & Left(partNum, 3) & "00\" & partNum & "\"
        strFilePath = prtFilePath & "Documents"
        
        If FolderExists(strFilePath) = True Then openDocumentHistoryFolder = strFilePath
    Case Else
        thousZeros = Left(partNum, 2) & "000\"
        hundZeros = Left(partNum, 3) & "00\"
        mainPath = mainFolder("docHisSearch")
        prtFilePath = mainPath & thousZeros & hundZeros
        FolName = Dir(prtFilePath & partNum & "*", vbDirectory)
        strFilePath = prtFilePath & FolName
        
        If Len(partNum) = 5 Or Right(partNum, 1) = "P" Then
            If Len(FolName) <> 0 Then openDocumentHistoryFolder = strFilePath
        Else
            openDocumentHistoryFolder = mainPath
        End If
End Select

If openBool = False Then Exit Function

If openDocumentHistoryFolder = "" Then
    If userData("dept") = "Design" Then
        DoCmd.OpenForm "frmCreateDesignFolders"
    Else
        Call openPath(mainPath)
    End If
Else
    Call openPath(openDocumentHistoryFolder)
End If

Exit Function
Err_Handler:
    Call handleError("wdbDirectoryFunctions", "openDocumentHistoryFolder", Err.DESCRIPTION, Err.number, CStr(partNum))
End Function

Function openModelV5Folder(partNumOriginal, Optional openFold As Boolean = True) As String
On Error GoTo Err_Handler

openModelV5Folder = ""

Dim partNum, thousZeros, hundZeros, FolName, mainfolderpath, strFilePath, prtpath, dPath

partNum = partNumOriginal & "_"
If partNum Like "D*" Then
    If openFold Then Call checkMkDir(mainFolder("ModelV5D"), Left(partNum, Len(partNum) - 1), "*")
    GoTo exit_handler
End If

If Left(partNum, 8) Like "[A-Z][A-Z]##[A-Z]##[A-Z]" Or Left(partNum, 7) Like "[A-Z][A-Z]##[A-Z]##" Or Left(partNum, 5) Like "##[A-Z]##" Then
    '---NCM PART NUMBER---
    'Examples: AB11A76A or AB11A76 or 11A76
    partNum = partNumOriginal
    If Not partNum Like "##[A-Z]##" Then partNum = Mid(partNum, 3, 5)
    
    mainfolderpath = mainFolder("ncmDrawingMaster")
    prtpath = mainfolderpath & Left(partNum, 3) & "00\" & partNum & "\"
    strFilePath = prtpath & "CATIA"
    
    If FolderExists(strFilePath) Then
        openModelV5Folder = strFilePath
        If openFold Then Call openPath(strFilePath)
    Else
        If openFold Then DoCmd.OpenForm "frmCreateDesignFolders"
    End If
Else
    '---NAM PART NUMBER---
    thousZeros = Left(partNum, 2) & "000\"
    hundZeros = Left(partNum, 3) & "00\"
    mainfolderpath = mainFolder("modelV5search")
    prtpath = mainfolderpath & thousZeros & hundZeros
tryAgain:
    FolName = Dir(prtpath & partNum & "*", vbDirectory)
    strFilePath = prtpath & FolName
    
    If Len(partNumOriginal) = 5 Or partNumOriginal Like "*P" Then
        If Len(FolName) = 0 Then
            If partNum Like "*_" Then
                partNum = Left(partNum, 5)
                GoTo tryAgain
            End If
            If openFold Then DoCmd.OpenForm "frmCreateDesignFolders"
        Else
            openModelV5Folder = strFilePath
            If openFold Then Call openPath(strFilePath)
        End If
    Else
        If openFold Then Call openPath(mainfolderpath)
    End If
End If

exit_handler:

Exit Function
Err_Handler:
    Call handleError("wdbDirectoryFunctions", "openModelV5Folder", Err.DESCRIPTION, Err.number)
End Function