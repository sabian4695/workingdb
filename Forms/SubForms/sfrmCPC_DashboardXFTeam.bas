Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Department_AfterUpdate()
On Error GoTo Err_Handler

Call registerCPCUpdates("tblCPC_XFTeams", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmCPC_Dashboard.ID)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

exit_handler:
    Exit Sub

Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
    Resume exit_handler

End Sub

Private Sub Form_Dirty(Cancel As Integer)

On Error GoTo Err_Handler

Me.Parent.lastModified = Now

exit_handler:
    Exit Sub

Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
    Resume exit_handler

End Sub

Private Sub memberName_AfterUpdate()
On Error GoTo Err_Handler

Me.Form.memberEmail = Me.memberName.column(2)

Call registerCPCUpdates("tblCPC_XFTeams", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmCPC_Dashboard.ID)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnDeleteMember_Click()

On Error GoTo Err_Handler

Dim ID As Long

If IsNull(Me.ID) Then
    Exit Sub
End If

If MsgBox("Are you sure you want to delete this member?", vbYesNo, "Warning") = vbYes Then
    ID = Me.ID
    Call registerCPCUpdates("tblCPC_XFTeams", ID, Me.memberName.name, Me.memberName, "Deleted", Form_frmCPC_Dashboard.ID)
    DoCmd.GoToRecord , , acNewRec
    dbExecute ("DELETE * FROM tblCPC_XFTeams WHERE [id] = " & ID)
    Me.Requery
End If

exit_handler:
    Exit Sub

Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
    Resume exit_handler
End Sub
