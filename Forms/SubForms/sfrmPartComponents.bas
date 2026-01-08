Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub componentNumber_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartComponents", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmPartAssemblyInfo.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Detail_Paint()
On Error Resume Next

Me.details.Transparent = Not DCount("recordId", "tblPartProject", "partNumber = '" & Me.componentNumber & "'") > 0

Me.remove.Transparent = False
If IsNull(Me.recordId) Then
    Me.details.Transparent = True
    Me.remove.Transparent = True
End If

End Sub

Private Sub details_Click()
On Error GoTo Err_Handler

openPartProject (Me.componentNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub finishLocator_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartComponents", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmPartAssemblyInfo.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub finishSubInv_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartComponents", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmPartAssemblyInfo.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_AfterInsert()
On Error GoTo Err_Handler

If Nz(Me.assemblyNumber, "") = "" Then
    Me.assemblyNumber = Form_frmPartAssemblyInfo.lblPartNumber.Caption
End If

Call registerPartUpdates("tblPartComponents", Me.recordId, Me.ActiveControl.name, Nz(Me.componentNumber), "Component Added", Form_frmPartAssemblyInfo.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub quantity_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartComponents", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmPartAssemblyInfo.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub remove_Click()
On Error GoTo Err_Handler

If IsNull(Me.recordId) Then Exit Sub

If (restrict(Environ("username"), TempVars!projectOwner) = True) Then
    MsgBox "Only project/service Engineers can do this", vbCritical, "Denied"
    Exit Sub
End If

If MsgBox("Are you sure you want to delete this?", vbYesNo, "Please confirm") <> vbYes Then Exit Sub

Call registerPartUpdates("tblPartComponents", Me.recordId, "Part Component", Me.componentNumber, "Deleted", Form_frmPartAssemblyInfo.lblPartNumber.Caption, Me.name)

dbExecute ("DELETE FROM tblPartComponents WHERE [recordId] = " & Me.recordId)

Me.Requery

MsgBox "Component Deleted", vbOKOnly, "Deleted"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
