Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Function validate() As String
validate = ""

Select Case True
    Case Nz(Me.capName) = ""
        validate = "Caption"
    Case Nz(Me.Link) = ""
        validate = "Link"
End Select

End Function

Private Sub btnSave_Click()
On Error GoTo Err_Handler

Dim val As String
val = validate

If val <> "" Then
    MsgBox "Please enter a " & val, vbCritical, "Fix this"
    Exit Sub
End If

If IsNull(DLookup("[Caption]", "[tblUserButtons]", "[ID] = " & Me.ID)) Then
    Me.User = Environ("username")
    Me.ButtonNum = Split(Me.buttonNumber, " ")(1)
End If

If Me.Dirty = True Then Me.Dirty = False

If InStr(Me.Link, "'") Or InStr(Me.Link, """") Then
    Me.Link = Replace(Me.Link, "'", "''")
    Me.Link = Replace(Me.Link, """", """""")
End If

dbExecute ("UPDATE tblUserButtons SET [tblUserButtons].Link = '" & Me.Link & "' WHERE [tblUserButtons].ID = " & Me.ID)

stopit:
If Me.Dirty = True Then Me.Undo

DoCmd.CLOSE
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub buttonNumber_AfterUpdate()
On Error GoTo Err_Handler

Me.filter = "[User] = '" & Environ("username") & "' AND [ButtonNum] = '" & Split(Me.buttonNumber, " ")(1) & "'"
Me.FilterOn = True

On Error GoTo errorCatch
Me.Link = Nz(DLookup("[Link]", "[tblUserButtons]", "[ID] = " & Me.ID))
Exit Sub
errorCatch:
Me.Link = ""
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.filter = "[User] = '" & Environ("username") & "' AND [ButtonNum] = '1'"
Me.FilterOn = True

On Error GoTo errorCatch
Me.Link = Nz(DLookup("[Link]", "[tblUserButtons]", "[ID] = " & Me.ID))
Exit Sub
errorCatch:
Me.Link = ""
Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler
Call Form_DASHBOARD.loadUserBtns

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Unload", Err.DESCRIPTION, Err.number)
End Sub

Private Sub reset_Click()
On Error GoTo Err_Handler

If IsNull(Me.ID) Then
    DoCmd.CLOSE
    Exit Sub
End If

dbExecute "DELETE FROM tblUserButtons WHERE [User] = '" & Environ("username") & "' And [ButtonNum] = '" & Me.ButtonNum & "' "

DoCmd.CLOSE

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
