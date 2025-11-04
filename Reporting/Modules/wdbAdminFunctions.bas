Option Compare Database
Option Explicit

Global Const SW_HIDE = 0
Global Const SW_SHOWNORMAL = 1
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWMAXIMIZED = 3

Private Type Rect
x1 As Long
y1 As Long
x2 As Long
y2 As Long
End Type

Private Declare PtrSafe Function apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" _
(ByVal hwnd As Long, r As Rect) As Long
Private Declare PtrSafe Function IsZoomed Lib "user32" _
(ByVal hwnd As Long) As Long
Private Declare PtrSafe Function MoveWindow Lib "user32" _
(ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, _
ByVal dx As Long, ByVal dy As Long, ByVal fRepaint As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" _
(ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Function privilege(pref, userInput) As Boolean
    privilege = DLookup("[" & pref & "]", "[tblPermissions]", "[User] = '" & userInput & "'")
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

Function grabVersion() As String
    grabVersion = DLookup("[Release]", "tblDBinfo", "[ID] = 1")
End Function