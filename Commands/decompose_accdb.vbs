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
		if Left(myObj.fullName ,1) = "s" Then
			oApplication.SaveAsText acForm, myObj.fullName, sExportpath & "\Forms\SubForms\" & myObj.fullName & ".form"
		Else
			oApplication.SaveAsText acForm, myObj.fullName, sExportpath & "\Forms\" & myObj.fullName & ".form"
		End
        oApplication.DoCmd.Close acForm, myObj.fullName
    Next
	
    For Each myObj In oApplication.CurrentProject.AllModules
        WScript.Echo "  " & myObj.fullName
		oApplication.SaveAsText acForm, myObj.fullName, sExportpath & "\Modules\" & myObj.fullName & ".bas"
    Next
	
    For Each myObj In oApplication.CurrentProject.AllMacros
        WScript.Echo "  " & myObj.fullName
        oApplication.SaveAsText acMacro, myObj.fullName, sExportpath & "\Macros\" & myObj.fullName & ".mod"
    Next
	
    For Each myObj In oApplication.CurrentProject.AllReports
        WScript.Echo "  " & myObj.fullName
		if Left(myObj.fullName ,1) = "s" Then
			oApplication.SaveAsText acForm, myObj.fullName, sExportpath & "\Reports\SubReports\" & myObj.fullName & ".rpt"
		Else
			oApplication.SaveAsText acForm, myObj.fullName, sExportpath & "\Reports\" & myObj.fullName & ".rpt"
		End
    Next
	
    For Each myObj In oApplication.CurrentDb.QueryDefs
        If Not Left(myObj.name, 3) = "~sq" Then 'exclude queries defined by the forms. Already included in the form itself
            WScript.Echo "  " & myObj.name
			if Left(myObj.fullName ,1) = "s" Then
				oApplication.SaveAsText acForm, myObj.fullName, sExportpath & "\Queries\SubQueries\" & myObj.fullName & ".qry"
			Else
				oApplication.SaveAsText acForm, myObj.fullName, sExportpath & "\Queries\" & myObj.fullName & ".qry"
			End
            oApplication.DoCmd.Close acQuery, myObj.name
        End If
    Next

    oApplication.CloseCurrentDatabase
    oApplication.Quit

    fso.DeleteFile sStubADPFilename

msgbox "Files Decomposed from " & sADPFilename, vbInformation, "Nicely Done"

End Function

Public Function getErr()
    Dim strError
    strError = vbCrLf & "----------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf & _
               "From " & Err.source & ":" & vbCrLf & _
               "    Description: " & Err.description & vbCrLf & _
               "    Code: " & Err.number & vbCrLf
    getErr = strError
End Function