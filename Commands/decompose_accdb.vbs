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
	
	Dim dbT, accT
	set accT = CreateObject("Access.Application")
	set dbT = accT.DBEngine.OpenDatabase(sStubADPFilename, False, False)

	dbT.Properties("AllowByPassKey") = True
	dbT.Close
	Set dbT = Nothing
	
    Dim oApplication
    Set oApplication = CreateObject("Access.Application")
    WScript.Echo "opening " & sStubADPFilename & " ..."
    oApplication.OpenCurrentDatabase sStubADPFilename
    oApplication.Visible = False

    Dim dctDelete
    Set dctDelete = CreateObject("Scripting.Dictionary")
    WScript.Echo "exporting..."
    Dim myObj
		
	Dim delFold
	Dim delFile

'delete all files
	If fso.FolderExists(sExportpath & "\Forms\") Then
		Set delFold = fso.GetFolder( sExportpath & "\Forms\")
		For Each delFile In delFold.Files
			WScript.Echo "  " & delFile.Path
			fso.DeleteFile delFile.Path, True ' True for force deletion
		Next
	end if
	
	If fso.FolderExists(sExportpath & "\Forms\SubForms\") Then
		Set delFold = fso.GetFolder( sExportpath & "\Forms\SubForms\")
		For Each delFile In delFold.Files
			WScript.Echo "  " & delFile.Path
			fso.DeleteFile delFile.Path, True ' True for force deletion
		Next
	end if
	
	If fso.FolderExists(sExportpath & "\Modules\") Then
		Set delFold = fso.GetFolder( sExportpath & "\Modules\")
		For Each delFile In delFold.Files
			WScript.Echo "  " & delFile.Path
			fso.DeleteFile delFile.Path, True ' True for force deletion
		Next
	end if
	
	If fso.FolderExists(sExportpath & "\Macros\") Then
		Set delFold = fso.GetFolder( sExportpath & "\Macros\")
		For Each delFile In delFold.Files
			WScript.Echo "  " & delFile.Path
			fso.DeleteFile delFile.Path, True ' True for force deletion
		Next
	end if
	
	If fso.FolderExists(sExportpath & "\Reports\") Then
		Set delFold = fso.GetFolder( sExportpath & "\Reports\")
		For Each delFile In delFold.Files
			WScript.Echo "  " & delFile.Path
			fso.DeleteFile delFile.Path, True ' True for force deletion
		Next
	end if
	
	If fso.FolderExists(sExportpath & "\Reports\SubReports\") Then
		Set delFold = fso.GetFolder( sExportpath & "\Reports\SubReports\")
		For Each delFile In delFold.Files
			WScript.Echo "  " & delFile.Path
			fso.DeleteFile delFile.Path, True ' True for force deletion
		Next
	end if
	
	If fso.FolderExists(sExportpath & "\Queries\") Then
		Set delFold = fso.GetFolder( sExportpath & "\Queries\")
		For Each delFile In delFold.Files
			WScript.Echo "  " & delFile.Path
			fso.DeleteFile delFile.Path, True ' True for force deletion
		Next
	end if
	
	If fso.FolderExists(sExportpath & "\Queries\SubQueries\") Then
		Set delFold = fso.GetFolder( sExportpath & "\Queries\SubQueries\")
		For Each delFile In delFold.Files
			WScript.Echo "  " & delFile.Path
			fso.DeleteFile delFile.Path, True ' True for force deletion
		Next
	end if
	
	
	Set delFold = Nothing
	
	
	'---FORMS---
    For Each myObj In oApplication.CurrentProject.AllForms
        WScript.Echo "  " & myObj.fullName
		'move all new files
		If Left(myObj.fullName ,1) = "s" Then
			oApplication.SaveAsText acForm, myObj.fullName, sExportpath & "\Forms\SubForms\" & myObj.fullName & ".form"
		Else
			oApplication.SaveAsText acForm, myObj.fullName, sExportpath & "\Forms\" & myObj.fullName & ".form"
		End If
    Next
	
	'---MODULES---
    For Each myObj In oApplication.CurrentProject.AllModules
        WScript.Echo "  " & myObj.fullName
		oApplication.SaveAsText acModule, myObj.fullName, sExportpath & "\Modules\" & myObj.fullName & ".bas"
    Next
	
    For Each myObj In oApplication.CurrentProject.AllMacros
        WScript.Echo "  " & myObj.fullName	
        oApplication.SaveAsText acMacro, myObj.fullName, sExportpath & "\Macros\" & myObj.fullName & ".mod"
    Next
	
	'---REPORTS---
    For Each myObj In oApplication.CurrentProject.AllReports
        WScript.Echo "  " & myObj.fullName
		If Left(myObj.fullName ,1) = "s" Then
			oApplication.SaveAsText acReport, myObj.fullName, sExportpath & "\Reports\SubReports\" & myObj.fullName & ".rpt"
		Else
			oApplication.SaveAsText acReport, myObj.fullName, sExportpath & "\Reports\" & myObj.fullName & ".rpt"
		End If
    Next
	
    For Each myObj In oApplication.CurrentDb.QueryDefs
        If Not Left(myObj.name, 3) = "~sq" Then 'exclude queries defined by the forms. Already included in the form itself
            WScript.Echo "  " & myObj.name
			If Left(myObj.name ,1) = "s" Then
				oApplication.SaveAsText acQuery, myObj.name, sExportpath & "\Queries\SubQueries\" & myObj.name & ".qry"
			Else
				oApplication.SaveAsText acQuery, myObj.name, sExportpath & "\Queries\" & myObj.name & ".qry"
			End If
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