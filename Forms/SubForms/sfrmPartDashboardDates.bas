Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub actualDate_AfterUpdate()
On Error GoTo Err_Handler

If (restrict(Environ("username"), TempVars!projectOwner, "Supervisor", True) = True) Then
    Me.ActiveControl = Me.ActiveControl.OldValue
    Call snackBox("error", "No can do", "Only project/service Supervisors/Managers can edit this field", "frmPartDashboard")
    Exit Sub
End If

Call registerPartUpdates("tblPartGates", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.gateTitle, Me.projectId)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Current()
On Error GoTo Err_Handler

Me.txtCF = Me.recordId

If IsNull(Me.recordId) = False Then
    Form_sfrmPartDashboard.Visible = True
    Form_sfrmPartDashboard.filter = "partGateId = " & Me.recordId & " AND [status] <> 'Closed'"
    Form_sfrmPartDashboard.FilterOn = True
    Form_sfrmPartDashboard.showClosedToggle.Value = False
    Form_sfrmPartDashboard.dueDate.Visible = True
    Form_sfrmPartDashboard.lblDue.Caption = "Due Date"
    Form_sfrmPartDashboard.showMyStepsToggle = False
    Call allowStepEdit
Else
    Form.sfrmPartDashboard.Visible = False
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Current", Err.DESCRIPTION, Err.number)
End Sub

Public Sub allowStepEdit()
On Error GoTo Err_Handler

Dim msgCap As String
msgCap = "Steps: " & Left(Me.gateTitle, 2)
Form_sfrmPartDashboard.Form.allowEdits = True
Call Form_sfrmPartDashboard.sfrmPartDashLock(False)

'are there steps in a previous gate that are open? MUST FINISH THOSE FIRST
If Me.recordId > DMin("[partGateId]", "tblPartSteps", "partProjectId = " & Me.projectId & " AND [status] <> 'Closed'") Then
    msgCap = "Steps: " & Left(Me.gateTitle, 2) & " (LOCKED)"
    Call Form_sfrmPartDashboard.sfrmPartDashLock(True)
    GoTo setMsg
End If

setMsg:
Form_sfrmPartDashboard.lblSteps.Caption = msgCap

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub plannedDate_AfterUpdate()
On Error GoTo Err_Handler

If (IsNull(Me.ActiveControl.OldValue) = False) Then
    If (restrict(Environ("username"), TempVars!projectOwner, "Supervisor", True) = True) Then
        Me.ActiveControl = Me.ActiveControl.OldValue
        Call snackBox("error", "No can do", "Only project/service Supervisors or Managers can edit this field once enterred", "frmPartDashboard")
        Exit Sub
    End If
End If

Call registerPartUpdates("tblPartGates", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.gateTitle, Me.projectId)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
