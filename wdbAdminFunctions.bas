Option Compare Database
Option Explicit

Global Const SW_HIDE = 0
Global Const SW_SHOWNORMAL = 1
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWMAXIMIZED = 3

Private Type RECT
x1 As Long
y1 As Long
x2 As Long
y2 As Long
End Type

Private Declare PtrSafe Function apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As Long, r As RECT) As Long
Public Declare PtrSafe Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function moveWindow Lib "user32" Alias "MoveWindow" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal fRepaint As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'to move windows by clicking
Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Public Declare PtrSafe Function ReleaseCapture Lib "user32.dll" () As Long

Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Dim AppX As Long, AppY As Long, AppTop As Long, AppLeft As Long, WinRECT As RECT, APointAPI As POINTAPI

Sub AppWindowSelect()
On Error GoTo err_handler
    'select application window
    GetWindowRect Application.hWndAccessApp, WinRECT
    AppTop = WinRECT.y1
    AppLeft = WinRECT.x1
    GetCursorPos APointAPI
    AppX = APointAPI.X
    AppY = APointAPI.Y

Exit Sub
err_handler:
    Call handleError("wdbAdminFunctions", "AppWindowSelect", Err.DESCRIPTION, Err.number)
End Sub

Sub AppWindowMove()
On Error GoTo err_handler
    'drag application window
    GetCursorPos APointAPI
    SetWindowPos Application.hWndAccessApp, 0, AppLeft - (AppX - APointAPI.X), AppTop - (AppY - APointAPI.Y), _
        0, 0, &H4 + &H1
Exit Sub
err_handler:
    Call handleError("wdbAdminFunctions", "AppWindowMove", Err.DESCRIPTION, Err.number)
End Sub

Sub moveForm(frm As Form)
On Error GoTo err_handler

    ReleaseCapture
    SendMessage frm.hWnd, &HA1, &H2, 0

Exit Sub
err_handler:
    Call handleError("wdbAdminFunctions", "moveForm", Err.DESCRIPTION, Err.number)
End Sub

Function logClick(modName As String, formName As String, Optional dataTag0 = "", Optional dataTag1 = "")
On Error Resume Next

If (CurrentProject.Path = "H:\dev") Then Exit Function
If DLookup("paramVal", "tblDBinfoBE", "parameter = '" & "recordAnalytics'") = False Then Exit Function

Dim rs1 As Recordset
Set rs1 = CurrentDb().OpenRecordset("tblAnalytics")

With rs1
    .addNew
        !Module = modName
        !Form = formName
        !userName = Environ("username")
        !dateUsed = Now()
        !dataTag0 = StrQuoteReplace(dataTag0)
        !dataTag1 = StrQuoteReplace(dataTag1)
    .Update
End With

rs1.Close
Set rs1 = Nothing

End Function

Function ap_DisableShift()

On Error GoTo errDisableShift
Dim db As DAO.Database
Dim prop As DAO.Property
Const conPropNotFound = 3270

Set db = CurrentDb()

db.Properties("AllowByPassKey") = False
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

Public Sub handleError(modName As String, activeCon As String, errDesc As String, errNum As Long)
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
    Case Else
        MsgBox errDesc, vbInformation, "Error Code: " & errNum
End Select

Dim strSQL As String

modName = StrQuoteReplace(modName)
errDesc = StrQuoteReplace(errDesc)
errNum = StrQuoteReplace(errNum)

strSQL = "INSERT INTO tblErrorLog(User,Form,Active_Control,Error_Date,Error_Description,Error_Number) VALUES ('" & _
 Environ("username") & "','" & modName & "','" & activeCon & "',#" & Now & "#,'" & errDesc & "'," & errNum & ")"

CurrentDb().Execute strSQL
End Sub

Function grabVersion() As String
On Error GoTo err_handler

Dim rs1 As Recordset
Set rs1 = CurrentDb().OpenRecordset("SELECT Release FROM tblDBinfo WHERE [ID] = 1", dbOpenSnapshot)
grabVersion = rs1!release
rs1.Close: Set rs1 = Nothing

Exit Function
err_handler:
    Call handleError("wdbAdminFunctions", "grabVersion", Err.DESCRIPTION, Err.number)
End Function

Function SixHatHideWindow(nCmdShow As Long)
On Error GoTo err_handler

    Dim loX As Long
    Dim loForm As Form
    On Error Resume Next
    Set loForm = Screen.ActiveForm

    If Err <> 0 Then
        loX = apiShowWindow(hWndAccessApp, nCmdShow)
        Err.clear
    End If

    If nCmdShow = SW_SHOWMINIMIZED And loForm.Modal = True Then
        MsgBox "Cannot minimize Access with " _
        & (loForm.Caption + " ") _
        & "form on screen"
    ElseIf nCmdShow = SW_HIDE And loForm.PopUp <> True Then
        MsgBox "Cannot hide Access with " _
        & (loForm.Caption + " ") _
        & "form on screen"
    Else
        loX = apiShowWindow(hWndAccessApp, nCmdShow)
    End If
    SixHatHideWindow = (loX <> 0)

Exit Function
err_handler:
    Call handleError("wdbAdminFunctions", "SixHatHideWindow", Err.DESCRIPTION, Err.number)
End Function

Sub SizeAccess(ByVal dx As Long, ByVal dy As Long)
On Error GoTo err_handler
'Set size of Access and center on Desktop

Const SW_RESTORE As Long = 9
Dim h As Long
Dim r As RECT
'
On Error Resume Next
'
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