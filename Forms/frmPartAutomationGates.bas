Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function applyTheFilters()
On Error GoTo Err_Handler
Dim filt

filt = ""

If Me.fltPartNumber <> "" Then
    If filt <> "" Then filt = filt & " AND "
    filt = filt & "partNumber = '" & Me.fltPartNumber & "'"
    Me.fltByUser = Null
    Me.fltModel = Null
    GoTo filtNow
End If

If Me.fltByUser <> "" Then filt = "(partNumber IN (SELECT partNumber FROM tblPartTeam WHERE person = '" & Me.fltByUser & "'))"

If Me.fltModel <> "" Then
    If filt <> "" Then filt = filt & " AND "
    filt = filt & "partNumber IN (SELECT partNumber FROM tblPartInfo WHERE programId = " & Me.fltModel & ")"
End If

filtNow:
Me.filter = filt
Me.FilterOn = filt <> ""

Exit Function
Err_Handler:
    Call handleError(Me.name, "applyTheFilters", Err.DESCRIPTION, Err.number)
End Function

Private Sub actualDate_AfterUpdate()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False

If Me.gateStatus <> 3 And IsNull(Me.actualDate) = False Then
    Me.gateStatus = 3
    Call gateStatus_AfterUpdate
End If

If Me.gateStatus <> 3 Then Exit Sub 'that means the change wasn't accepted!

Call registerPartUpdates("tblPartAssemblyGates", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.projectId)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub autoGateFiles_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmPartAttachments", , , "partNumber = '" & Me.partNumber & "' AND docTypeId = 23"
Form_frmPartAttachments.TpartNumber = Me.partNumber
Form_frmPartAttachments.TtestId = Me.recordId
Form_frmPartAttachments.TprojectId = DMax("recordId", "tblPartProject", "partNumber = '" & Me.partNumber & "'")
Form_frmPartAttachments.itemName.Caption = "Work Instructions"
Form_frmPartAttachments.secondaryType = 23

Form_frmPartAttachments.newAttachment.Visible = (DCount("recordId", "tblPartAttachmentsSP", "partNumber = '" & Me.partNumber & "' AND documentType = 23") = 0)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Detail_Paint()
On Error Resume Next
Me.autoGateFiles.Transparent = Me.templateGateId <> 362
End Sub

Private Sub details_Click()
On Error GoTo Err_Handler

openPartProject (Me.partNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltByUser_AfterUpdate()
On Error GoTo Err_Handler
applyTheFilters
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltModel_AfterUpdate()
On Error GoTo Err_Handler
applyTheFilters
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltPartNumber_AfterUpdate()
On Error GoTo Err_Handler
applyTheFilters
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim allowIt As Boolean
allowIt = Not restrict(Environ("username"), "Automation")

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub gateHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory", acNormal, , "[tableName] = 'tblPartAssemblyGates' AND [tableRecordId] = " & Me.recordId

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub gateNotes_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartAssemblyGates", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.projectId)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub gateStatus_AfterUpdate()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False

If Me.gateStatus = 3 Then 'marking as complete! need to check WI for bully test
    If (Me.templateGateId = 362) And (DCount("recordId", "tblPartAttachmentsSP", "testId = " & [recordId] & " AND documentType = 23") = 0) Then
        Call snackBox("error", "Not Yet!", "Please upload your Work Instructions first.", Me.name)
        Me.gateStatus = 2
        Me.actualDate = Null
        Exit Sub
    End If
    If IsNull(Me.actualDate) Then
        Me.actualDate = Date
        Call registerPartUpdates("tblPartAssemblyGates", Me.recordId, Me.actualDate.name, Me.actualDate.OldValue, Date, Me.partNumber, Me.projectId)
    End If
Else
    Me.actualDate = Null
    Call registerPartUpdates("tblPartAssemblyGates", Me.recordId, Me.actualDate.name, Me.actualDate.OldValue, Null, Me.partNumber, Me.projectId)
End If

Call registerPartUpdates("tblPartAssemblyGates", Me.recordId, Me.ActiveControl.name, Me.gateStatus.OldValue, Me.gateStatus.column(1), Me.partNumber, Me.projectId)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub partHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory", acNormal, , "[tableName] = 'tblPartAssemblyGates' AND [partNumber] = '" & Me.partNumber & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Public Sub refresh_Click()
On Error GoTo Err_Handler
Me.Requery
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
