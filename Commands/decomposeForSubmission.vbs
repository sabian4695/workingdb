Option Explicit

Const acForm = 2
Const acModule = 5
Const acMacro = 4
Const acReport = 3
Const acQuery = 1

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim stream
Set stream = CreateObject("ADODB.Stream")

Dim sADPFilename
sADPFilename = fso.GetAbsolutePathName(WScript.Arguments(0))

Dim sExportpath
sExportpath = fso.GetAbsolutePathName(WScript.Arguments(1))

exportModulesTxt sADPFilename, sExportpath

If (Err <> 0) And (Err.description <> Null) Then
    MsgBox Err.description, vbExclamation, "Error"
    Err.clear
End If

Function exportModulesTxt(sADPFilename, sExportpath)

    Dim myType, myName, myPath, sStubADPFilename
    myType = fso.GetExtensionName(sADPFilename)
    myName = fso.GetBaseName(sADPFilename)
    myPath = fso.GetParentFolderName(sADPFilename)

    sStubADPFilename = sExportpath & "\" & myName & "_stub." & myType

    WScript.Echo "copy stub to " & sStubADPFilename & "..."
    On Error Resume Next
        fso.CreateFolder (sExportpath)
    On Error GoTo 0
    fso.CopyFile sADPFilename, sStubADPFilename

    WScript.Echo "starting Access..."
    Dim oApplication
    Set oApplication = CreateObject("Access.Application")
    WScript.Echo "opening " & sStubADPFilename & " ..."
    oApplication.OpenCurrentDatabase sStubADPFilename
    oApplication.Visible = False

    Dim dctDelete
    Set dctDelete = CreateObject("Scripting.Dictionary")
    WScript.Echo "exporting..."
    Dim myObj

    For Each myObj In oApplication.CurrentProject.AllForms
        WScript.Echo "  " & myObj.fullName
        oApplication.SaveAsText acForm, myObj.fullName, sExportpath & "\" & myObj.fullName & ".form"
        oApplication.DoCmd.Close acForm, myObj.fullName
    Next
    For Each myObj In oApplication.CurrentProject.AllModules
        WScript.Echo "  " & myObj.fullName
        oApplication.SaveAsText acModule, myObj.fullName, sExportpath & "\" & myObj.fullName & ".bas"
    Next
    For Each myObj In oApplication.CurrentProject.AllMacros
        WScript.Echo "  " & myObj.fullName
        oApplication.SaveAsText acMacro, myObj.fullName, sExportpath & "\" & myObj.fullName & ".mod"
    Next
    For Each myObj In oApplication.CurrentProject.AllReports
        WScript.Echo "  " & myObj.fullName
        oApplication.SaveAsText acReport, myObj.fullName, sExportpath & "\" & myObj.fullName & ".rpt"
    Next
    For Each myObj In oApplication.CurrentDb.QueryDefs
        If Not Left(myObj.name, 3) = "~sq" Then 'exclude queries defined by the forms. Already included in the form itself
            WScript.Echo "  " & myObj.name
            oApplication.SaveAsText acQuery, myObj.name, sExportpath & "\" & myObj.name & ".qry"
            oApplication.DoCmd.Close acQuery, myObj.name
        End If
    Next

    oApplication.CloseCurrentDatabase
    oApplication.Quit

    fso.DeleteFile sStubADPFilename

msgbox "Dev submitted successfully", vbInformation, "Nicely Done"

dim strUser
strUser = CreateObject("WScript.Network").UserName
fso.deletefile "H:\dev\WorkingDB_" & strUser & "_dev.accdb"

End Function

Public Function getErr()
    Dim strError
    strError = vbCrLf & "----------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf & _
               "From " & Err.source & ":" & vbCrLf & _
               "    Description: " & Err.description & vbCrLf & _
               "    Code: " & Err.number & vbCrLf
    getErr = strError
End Function