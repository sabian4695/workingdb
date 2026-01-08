Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnSave_Click()
On Error GoTo Err_Handler

Dim oldV, newV

If Me.Dirty Then
    On Error Resume Next
    oldV = Nz(DLookup("stepActionTitle", "tblPartStepActions", "recordId = " & Me.stepActionId.OldValue), "")
    newV = Nz(DLookup("stepActionTitle", "tblPartStepActions", "recordId = " & Me.stepActionId), "")
    On Error GoTo Err_Handler
    
    If oldV <> newV Then
        Call registerPartUpdates("tblPartSteps", Me.recordId, Me.stepActionId.name, oldV, newV, Form_frmPartDashboard.partNumber, , Form_frmPartDashboard.recordId)
    End If
    If Me.stepType <> Me.stepType.OldValue Then
        Call registerPartUpdates("tblPartSteps", Me.recordId, Me.stepType.name, Me.stepType.OldValue, Me.stepType, Form_frmPartDashboard.partNumber, , Form_frmPartDashboard.recordId)
    End If
    
End If

DoCmd.CLOSE
Form_frmPartDashboard.partDash_refresh_Click

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub deletebtn_Click()
On Error GoTo Err_Handler

If MsgBox("Are you sure you want to delete this entire step?", vbYesNo, "Please confirm") <> vbYes Then Exit Sub

Call registerPartUpdates("tblPartSteps", Me.recordId, "Step", Me.stepType, "Deleted", Form_frmPartDashboard.partNumber, , Form_frmPartDashboard.recordId)

Dim db As Database
Set db = CurrentDb()
db.Execute "DELETE FROM tblPartTrackingApprovals WHERE tableName = 'tblPartSteps' AND tableRecordId = " & Me.recordId
db.Execute "DELETE FROM tblPartSteps WHERE recordId = " & Me.recordId
DoCmd.CLOSE
Form_frmPartDashboard.partDash_refresh_Click

Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub moveDown_Click()
On Error GoTo Err_Handler

If Me.indexOrder = DMax("indexOrder", "tblPartSteps", "partGateId = " & Me.partGateId) Then Exit Sub

Dim oldIndex, newIndex
oldIndex = Me.indexOrder
newIndex = oldIndex + 1

dbExecute "UPDATE tblPartSteps SET indexOrder = " & oldIndex & " WHERE partGateId = " & Me.partGateId & " AND indexOrder = " & newIndex

Me.indexOrder = newIndex
Me.Dirty = False

Call registerPartUpdates("tblPartSteps", Me.recordId, "indexOrder", oldIndex, newIndex, Form_frmPartDashboard.partNumber, , Form_frmPartDashboard.recordId)

DoCmd.CLOSE
Form_frmPartDashboard.partDash_refresh_Click

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub moveUp_Click()
On Error GoTo Err_Handler

If Me.indexOrder = 1 Then Exit Sub
Dim oldIndex, newIndex
oldIndex = Me.indexOrder
newIndex = oldIndex - 1

dbExecute "UPDATE tblPartSteps SET indexOrder = " & oldIndex & " WHERE partGateId = " & Me.partGateId & " AND indexOrder = " & newIndex

Me.indexOrder = newIndex
Me.Dirty = False

Call registerPartUpdates("tblPartSteps", Me.recordId, "indexOrder", oldIndex, newIndex, Form_frmPartDashboard.partNumber, , Form_frmPartDashboard.recordId)

DoCmd.CLOSE
Form_frmPartDashboard.partDash_refresh_Click

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
