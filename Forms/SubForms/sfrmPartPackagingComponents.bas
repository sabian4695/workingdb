Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub componentPN_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartPackagingComponents", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmPartInformation.partNumber, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub componentQuantity_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartPackagingComponents", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmPartInformation.partNumber, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub componentType_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartPackagingComponents", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmPartInformation.partNumber, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub remove_Click()
On Error GoTo Err_Handler

If IsNull(Me.recordId) Then Exit Sub

If (Not restrict(Environ("username"), "Packaging") Or Not restrict(Environ("username"), "Project")) = False Then
    MsgBox "Only project/service Engineers can do this", vbCritical, "Denied"
    Exit Sub
End If

If MsgBox("Are you sure you want to delete this?", vbYesNo, "Please confirm") <> vbYes Then Exit Sub

Call registerPartUpdates("tblPartPackagingInfo", Me.recordId, "Part Packaging Info", Me.componentType.column(1), "Deleted", Form_frmPartInformation.partNumber, Me.name)

dbExecute ("DELETE FROM tblPartPackagingComponents WHERE [recordId] = " & Me.recordId)
Me.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
