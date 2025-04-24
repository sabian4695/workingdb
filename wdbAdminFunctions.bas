Option Compare Database
Option Explicit

Global Const SW_HIDE = 0
Global Const SW_SHOWNORMAL = 1
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWMAXIMIZED = 3
Global Const SW_RESTORE = 9

Private Type RECT
x1 As Long
y1 As Long
x2 As Long
y2 As Long
End Type

Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As Long, r As RECT) As Long
Public Declare PtrSafe Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function moveWindow Lib "user32" Alias "MoveWindow" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal fRepaint As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Dim AppX As Long, AppY As Long, AppTop As Long, AppLeft As Long, WinRECT As RECT

Function readyForPublish() As Boolean
On Error GoTo err_handler

readyForPublish = False

'First, try to compile
Dim compileMe As Object
Set compileMe = Application.VBE.CommandBars.FindControl(type:=msoControlButton, ID:=578)

'try to compile
If compileMe.Enabled Then compileMe.Execute

'do a double check
If (Application.IsCompiled = False) Then
    If compileMe.Enabled Then compileMe.Execute
End If
DoEvents

'--Can you even do this?--
Dim errorMsg As String: errorMsg = ""
If (Application.IsCompiled = False) Then errorMsg = "Please compile codebase"
If Not ((Environ("username") <> "brownj") Or (Environ("username") <> "georgemi")) Then errorMsg = "You must be an owner to do that"

If errorMsg <> "" Then
    MsgBox errorMsg, vbCritical, "Access Denied"
    Exit Function
End If

readyForPublish = True

Exit Function
err_handler:
    Call handleError("wdbAdminFunctions", "readyForPublish", Err.DESCRIPTION, Err.number)
End Function

Function logClick(modName As String, formName As String, Optional dataTag0 = "")
On Error Resume Next

If DLookup("paramVal", "tblDBinfoBE", "parameter = '" & "recordAnalytics'") = False Then Exit Function

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("tblAnalytics")

With rs1
    .addNew
        !module = modName
        !Form = formName
        !userName = Environ("username")
        !dateUsed = Now()
        !dataTag0 = StrQuoteReplace(dataTag0)
        !dataTag1 = TempVars!wdbVersion
    .Update
End With

rs1.Close
Set rs1 = Nothing
Set db = Nothing

End Function

Function ap_DisableShift()

On Error GoTo errDisableShift
Dim db As DAO.Database
Dim prop As DAO.Property
Const conPropNotFound = 3270

Set db = CurrentDb()

db.Properties("AllowByPassKey") = False
Set db = Nothing
Exit Function

errDisableShift:
If Err = conPropNotFound Then
    Set prop = db.CreateProperty("AllowByPassKey", dbBoolean, False)
    db.Properties.Append prop
    Resume Next
    Else
    MsgBox "Function 'ap_DisableShift' did not complete successfully."
    Exit Function
End If

End Function

Function ap_EnableShift()

On Error GoTo errEnableShift
Dim db As DAO.Database
Dim prop As DAO.Property
Const conPropNotFound = 3270

Set db = CurrentDb()
db.Properties("AllowByPassKey") = True
Set db = Nothing
Exit Function

errEnableShift:
If Err = conPropNotFound Then
Set prop = db.CreateProperty("AllowByPassKey", dbBoolean, True)
db.Properties.Append prop
Resume Next
Else
MsgBox "Function 'ap_DisableShift' did not complete successfully."
Exit Function
End If

End Function

Public Sub handleError(modName As String, activeCon As String, errDesc As String, errNum As Long, Optional dataTag As String = "")
On Error Resume Next
If (CurrentProject.Path = "H:\dev") Then
    MsgBox errDesc, vbInformation, "Error Code: " & errNum
    Exit Sub
End If

Select Case errNum
    Case 53
        MsgBox "File Not Found", vbInformation, "Error Code: " & errNum
        Exit Sub
    Case 3011
        MsgBox "Looks like I'm having issues connecting to SharePoint. Please reopen when you can", vbInformation, "Error Code: " & errNum
    Case 490, 52, 75
        MsgBox "I cannot open this file or location - check if it has been moved or deleted. Or - you do not have proper access to this location", vbInformation, "Error Code: " & errNum
        Exit Sub
    Case 3022
        MsgBox "A record with this key already exists. I cannot create another!", vbInformation, "Error Code: " & errNum
    Case 3167
        MsgBox "Looks like you already deleted that record", vbInformation, "Error Code: " & errNum
        Exit Sub
    Case 94
        MsgBox "Hmm. Looks like something is missing. Check for an empty field", vbInformation, "Error Code: " & errNum
    Case 3151
        MsgBox "You're not connected to Oracle. Just FYI, Oracle connection does not work outside of VMWare.", vbInformation, "Error Code: " & errNum
        Exit Sub
    Case 429
        If modName = "frmCatiaMacros" Then
            MsgBox "Looks like Catia isn't open", vbInformation, "Error Code: " & errNum
            Exit Sub
        Else
            MsgBox errDesc, vbInformation, "Error Code: " & errNum
        End If
    Case 3343
        MsgBox "Error. Please re-open WorkingDB to reset.", vbCritical, "Error Code: " & errNum
    Case Else
        MsgBox errDesc, vbInformation, "Error Code: " & errNum
End Select

Dim strSQL As String

modName = StrQuoteReplace(modName)
errDesc = StrQuoteReplace(errDesc)
errNum = StrQuoteReplace(errNum)
dataTag = StrQuoteReplace(dataTag)

strSQL = "INSERT INTO tblErrorLog(User,Form,Active_Control,Error_Date,Error_Description,Error_Number,databaseVersion,dataTag0) VALUES ('" & _
 Environ("username") & "','" & modName & "','" & activeCon & "',#" & Now & "#,'" & errDesc & "'," & errNum & ",'" & TempVars!wdbVersion & "','" & dataTag & "')"

dbExecute strSQL
End Sub

Function grabVersion() As String
On Error GoTo err_handler

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("SELECT Release FROM tblDBinfo WHERE [ID] = 1", dbOpenSnapshot)
grabVersion = rs1!release
rs1.Close: Set rs1 = Nothing
Set db = Nothing

Exit Function
err_handler:
    Call handleError("wdbAdminFunctions", "grabVersion", Err.DESCRIPTION, Err.number)
End Function

Sub maximizeAccess()
On Error GoTo err_handler

Dim h As Long
Dim r As RECT

'On Error Resume Next

h = Application.hWndAccessApp
'If maximised, restore
If (IsZoomed(h) = False) Then ShowWindow h, SW_SHOWMAXIMIZED

Exit Sub
err_handler:
    Call handleError("wdbAdminFunctions", "SizeAccess", Err.DESCRIPTION, Err.number)
End Sub

Sub SizeAccess(ByVal dx As Long, ByVal dy As Long)
On Error GoTo err_handler
'Set size of Access and center on Desktop

Dim h As Long
Dim r As RECT

On Error Resume Next

h = Application.hWndAccessApp
'If maximised, restore
If (IsZoomed(h)) Then ShowWindow h, SW_RESTORE
'
'Get available Desktop size
GetWindowRect GetDesktopWindow(), r
If ((r.x2 - r.x1) - dx) < 0 Or ((r.y2 - r.y1) - dy) < 0 Then
'Desktop smaller than requested size
'so size to Desktop
moveWindow h, r.x1, r.y1, r.x2, r.y2, True
Else
'Adjust to requested size and center
moveWindow h, _
r.x1 + ((r.x2 - r.x1) - dx) \ 2, _
r.y1 + ((r.y2 - r.y1) - dy) \ 2, _
dx, dy, True
End If

Exit Sub
err_handler:
    Call handleError("wdbAdminFunctions", "SizeAccess", Err.DESCRIPTION, Err.number)
End Sub