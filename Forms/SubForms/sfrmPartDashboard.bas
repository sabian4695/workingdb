Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub approvals_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmPartTrackingApprovals", , , "[tableRecordId] = " & Me.recordId & " AND [tableName] = 'tblPartSteps'"

Form_frmPartTrackingApprovals.TpartNumber = Me.partNumber
Form_frmPartTrackingApprovals.TtableName = "tblPartSteps"
Form_frmPartTrackingApprovals.TtableRecordId = Me.recordId
Form_frmPartTrackingApprovals.itemName.Caption = Me.stepType
Form_frmPartTrackingApprovals.approvalNote.tag = Nz(Me.responsible)

Dim allowIt As Boolean
allowIt = Me.status <> "Closed"

Form_frmPartTrackingApprovals.approve.Enabled = allowIt
Form_frmPartTrackingApprovals.allowEdits = allowIt
Form_frmPartTrackingApprovals.sendFeedback.Enabled = allowIt
Form_frmPartTrackingApprovals.nudge.Enabled = allowIt
Form_frmPartTrackingApprovals.remove.Enabled = allowIt
Form_frmPartTrackingApprovals.newApproval.Enabled = allowIt
Form_frmPartTrackingApprovals.save.Enabled = allowIt

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnEditInfo_Click()
On Error GoTo Err_Handler

If restrict(Environ("username"), TempVars!projectOwner, "Supervisor", True) Then
    Call snackBox("error", "You can't edit a closed step", "Only a project/service manager can edit a step - looks like that's not you, sorry.", "frmPartDashboard")
    Exit Sub
End If

DoCmd.OpenForm "frmPartStepEdit", acNormal, , "recordId = " & Me.recordId

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub closeStep_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False
If closeProjectStep(Me.recordId, "frmPartDashboard") Then
    Call snackBox("success", "Success!", "Step Closed", "frmPartDashboard")
    Me.Requery
    Me.refresh
    Call updateLastUpdate
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub description_AfterUpdate()
On Error GoTo Err_Handler

If IsNull(Me.closeDate) = False Then
    If restrict(Environ("username"), TempVars!projectOwner, "Manager") = True Then
        Me.ActiveControl = Me.ActiveControl.OldValue
        Call snackBox("error", "You can't edit a closed step", "Only a project/service manager and edit a closed step - looks like that's not you, sorry.", "frmPartDashboard")
        Exit Sub
    End If
End If

Call registerPartUpdates("tblPartSteps", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.type, Me.partProjectId)
Call updateLastUpdate

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Detail_Paint()
On Error Resume Next

Me.closeStep.Transparent = Nz(Me.stepAction, "") = "closeStep"

End Sub

Private Sub dueDate_AfterUpdate()
On Error GoTo Err_Handler

If IsNull(Me.closeDate) = False And restrict(Environ("username"), TempVars!projectOwner, "Manager") Then
    Me.dueDate = Me.dueDate.OldValue
    Call snackBox("error", "You can't edit a closed step", "Only a project/service manager and edit a closed step - looks like that's not you, sorry.", "frmPartDashboard")
    Exit Sub
End If

'restrict dueDates to be within GATE planned dates
Dim db As Database
Set db = CurrentDb()
Dim rsGates As Recordset
Dim startdate As Date, endDate As Date

Set rsGates = db.OpenRecordset("SELECT * from tblPartGates WHERE projectId = " & Me.partProjectId)

rsGates.FindFirst "gateTitle Like 'G1*'"

If rsGates.noMatch Then GoTo skipGateCheck
If rsGates!recordId = Me.partGateId Then 'this is the FIRST gate
    startdate = CDate("1/1/2000")
    endDate = Nz(rsGates!plannedDate, CDate("1/1/2000"))
Else
    rsGates.FindFirst "recordId = " & Me.partGateId
    endDate = Nz(rsGates!plannedDate, CDate("1/1/2000"))
    
tryAgain:
    rsGates.MovePrevious

    If rsGates.BOF Then GoTo stopThis
    If IsNull(rsGates!actualDate) Then GoTo tryAgain
    
stopThis:
    startdate = Nz(rsGates!actualDate, CDate("1/1/2000"))
End If

rsGates.CLOSE
Set rsGates = Nothing
Set db = Nothing

If endDate = CDate("1/1/2000") Then
    Me.dueDate = Me.dueDate.OldValue
    Call snackBox("error", "Missing Planned Date", "One or more of your gates needs a planned date added", "frmPartDashboard")
    Exit Sub
End If

If Not IsNull(Me.dueDate) Then
    If Not (Me.dueDate < endDate And Me.dueDate > startdate) Then 'if this is not within the gate dates, don't allow it!
        Me.dueDate = Me.dueDate.OldValue
        Call snackBox("error", "Outside of Gate Dates", "Due date must be between the dates for each gate", "frmPartDashboard")
        Exit Sub
    End If
End If

skipGateCheck:

Call registerPartUpdates("tblPartSteps", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.type, Me.partProjectId)

If Not IsNull(Me.dueDate.OldValue) Then
    If MsgBox("Do you want to adjust all future dates with this change?", vbYesNo, "Please confirm") = vbYes Then Call recalcStepDueDates(Me.partProjectId, Me.dueDate.OldValue, countWorkdays(Me.dueDate.OldValue, Me.dueDate))
End If

Me.Requery
Call updateLastUpdate

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub newStep_Click()
On Error GoTo Err_Handler

If IsNull(Me.closeDate) = False And restrict(Environ("username"), Nz(TempVars!projectOwner, ""), "Manager") Then
    Call snackBox("error", "You can't edit a closed gate", "Only a project/service manager and edit a closed gate - looks like that's not you, sorry.", "frmPartDashboard")
    Exit Sub
End If

Dim response
response = MsgBox("Click YES to add a Change Template, NO to add a step from scratch", vbYesNo, "Error")
If response = vbYes Then
    DoCmd.OpenForm "frmPartChangeTemplateSelector"
    
    Form_frmPartChangeTemplateSelector.partNumber = Form_frmPartDashboard.partNumber
    Form_frmPartChangeTemplateSelector.partProjectId = Me.partProjectId
    Form_frmPartChangeTemplateSelector.partGateId = Me.partGateId
ElseIf response = vbNo Then
    DoCmd.OpenForm "frmNewPartStep"

    Form_frmNewPartStep.partNumber = Form_frmPartDashboard.partNumber
    Form_frmNewPartStep.partProjectId = Me.partProjectId
    Form_frmNewPartStep.partGateId = Me.partGateId
    Form_frmNewPartStep.indexVal = Me.indexOrder
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub nudgeApprovers_Click()
On Error GoTo Err_Handler

If restrict(Environ("username"), TempVars!projectOwner) = False Then GoTo theCorrectFellow 'is the bro an owner?
If Me.responsible = userData("Dept") And DCount("recordId", "tblPartTeam", "person = '" & Environ("username") & "' AND partNumber = '" & Me.partNumber & "'") > 0 Then GoTo theCorrectFellow 'if the bro is responsible AND CHECK IF ON CF TEAM
If restrict(Environ("username"), Me.responsible, "Manager") = False Then GoTo theCorrectFellow  'is the bro a manager in the department of the "responsible" person?
MsgBox "Only the 'Responsible' person, their manager, or a project/service Engineer can nudge these approvers", vbCritical, "Woops"
Exit Sub
theCorrectFellow:

Dim db As Database
Set db = CurrentDb()
Dim rs3 As Recordset
Set rs3 = db.OpenRecordset("select * from tblPartTrackingApprovals where tableName = 'tblPartSteps' and approvedOn is null and tableRecordId = " & Me.recordId)

If rs3.RecordCount = 0 Then
    MsgBox "No open approvals found!", vbInformation, "Hmm.."
    Exit Sub
End If

Dim stepTitle As String
stepTitle = Me.stepType

Dim sendTo As String
Dim body As String
body = emailContentGen("You've been nudged...", "Nudge Notification", "You've been nudged by " & getFullName() & " to approve this step", stepTitle, "Part Number: " & Me.partNumber, "Requested By: " & rs3!requestedBy, "Requested On: " & CStr(rs3!requestedDate), appName:="Part Project", appId:=Me.partNumber)

Do While Not rs3.EOF 'loop through all the approvals
    sendTo = Nz(rs3!approver)
    If sendTo = Environ("username") Then GoTo nextOne
    
    If Nz(sendTo) = "" Then 'if approver not specified, look through cross functional team and see if anyone matches
        Dim rs1 As Recordset, rs2 As Recordset
        Set rs1 = db.OpenRecordset("select * from tblPartTeam where partNumber = '" & Me.partNumber & "'")
        Do While Not rs1.EOF
            Set rs2 = db.OpenRecordset("select * from tblPermissions where user = '" & rs1!person & "'")
            If rs2!dept <> rs3!dept Or rs2!Level <> rs3!reqLevel Then GoTo nextOne 'dept/level dont match. skip this person
            sendTo = rs2!User
            If sendTo = Environ("username") Then
                MsgBox "can't nudge yourself", vbInformation, "Darn"
                GoTo nextOne
            End If
            If sendNotification(sendTo, 1, 2, "Please approve step " & Me.type & " for " & Me.partNumber, body, "Part Project", Me.partNumber) = True Then
                MsgBox "Notification sent to " & sendTo & "!", vbInformation, "Well done."
                Call registerPartUpdates("tblPartSteps", Me.recordId, "Nudge", "From: " & Environ("username"), "To: " & sendTo, Me.partNumber, Me.type, Me.partProjectId)
            End If
nextOne:
            rs1.MoveNext
        Loop
        If sendTo = "" Then MsgBox rs3!dept & " " & rs3!reqLevel & " not found on the cross functional team", vbInformation, "Sorry, can't do that"
    Else
        If sendNotification(sendTo, 1, 2, "Please approve step " & Me.type & " for " & Me.partNumber, body, "Part Project", Me.partNumber) = True Then
            MsgBox "Notification sent to " & sendTo & "!", vbInformation, "Well done."
            Call registerPartUpdates("tblPartSteps", Me.recordId, "Nudge", "From: " & Environ("username"), "To: " & sendTo, Me.partNumber, Me.type, Me.partProjectId)
        End If
    End If
nextApproval:
    rs3.MoveNext
Loop

On Error Resume Next
rs1.CLOSE
Set rs1 = Nothing
rs2.CLOSE
Set rs2 = Nothing
rs3.CLOSE
Set rs3 = Nothing
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub nudgeResponsible_Click()
On Error GoTo Err_Handler

Dim sendTo As String, notiSent As Boolean, stepTitle As String
If Nz(Me.responsible) = "" Then
    Call snackBox("error", "No can do", "Need a 'responsible' department to nudge", "frmPartDashboard")
    Exit Sub
End If
notiSent = False

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset, rs2 As Recordset
Set rs1 = db.OpenRecordset("select * from tblPartTeam where partNumber = '" & Me.partNumber & "'")
Do While Not rs1.EOF
    Set rs2 = db.OpenRecordset("select * from tblPermissions where user = '" & rs1!person & "'")
    If rs2!dept <> Me.responsible Then GoTo nextOne 'if dept isn't the same, skip this user
    sendTo = rs2!User
    If sendTo = Environ("username") Then
        Call snackBox("error", "No can do", "You can't nudge yourself!", "frmPartDashboard")
        GoTo nextOne
    End If

    stepTitle = Me.stepType

    Dim body As String
    body = emailContentGen("You've been nudged...", "Nudge Notification", "You've been nudged by " & getFullName() & " to complete this step", stepTitle, "Part Number: " & Me.partNumber, "Requested By: " & Environ("username"), "Requested On: " & CStr(Date), appName:="Part Project", appId:=Me.partNumber)
    If sendNotification(sendTo, 1, 2, "Please complete step " & Me.type & " for " & Me.partNumber, body, "Part Project", Me.partNumber) = True Then
        Call snackBox("success", "Well done.", "Notification sent to " & sendTo & "!", "frmPartDashboard")
        Call registerPartUpdates("tblPartSteps", Me.recordId, "Nudge", "From: " & Environ("username"), "To: " & sendTo, Me.partNumber, Me.type, Me.partProjectId)
        notiSent = True
    End If
nextOne:
    rs1.MoveNext
Loop

If Not notiSent Then Call snackBox("error", "Woops", "No one found", "frmPartDashboard")

On Error Resume Next
rs1.CLOSE
Set rs1 = Nothing
rs2.CLOSE
Set rs2 = Nothing
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub showClosedToggle_Click()
On Error GoTo Err_Handler

Me.showMyStepsToggle = False

Dim extFilt As String
If Me.ActiveControl.Value Then
    extFilt = "="
    Me.lblDue.Caption = "Closed"
Else
    extFilt = "<>"
    Me.lblDue.Caption = "Due"
End If

Me.dueDate.Visible = Not Me.ActiveControl.Value
Me.filter = "partGateId = " & Form_sfrmPartDashboardDates.recordId & " AND [status] " & extFilt & " 'Closed'"
Me.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Sub updateLastUpdate()
On Error GoTo Err_Handler

If Me.Recordset.RecordCount = 0 Then Exit Sub

Me.lastUpdatedDate = Now()
Me.lastUpdatedBy = Environ("username")

If Me.status = "Not Started" Then
    Me.status = "In Progress"
    Call registerPartUpdates("tblPartSteps", Me.recordId, "Status", "Not Started", "In Progress", Me.partNumber, Me.stepType, Me.partProjectId)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, "updateLastUpdate", Err.DESCRIPTION, Err.number)
End Sub

Private Sub showMyStepsToggle_Click()
On Error GoTo Err_Handler

Me.showClosedToggle = False

Dim extFilt As String
extFilt = ""
If Me.ActiveControl.Value Then extFilt = " AND (recordId IN (Select tableRecordId FROM tblPartTrackingApprovals WHERE dept = '" & userData("Dept", Me.userSelect) & _
    "' AND reqLevel = '" & userData("Level", Me.userSelect) & "' AND approvedOn is null AND partNumber = '" & Me.partNumber & "') OR (responsible = '" & userData("Dept", Me.userSelect) & "'))"

Me.filter = "partGateId = " & Form_sfrmPartDashboardDates.recordId & " AND [status] <> 'Closed'" & extFilt
Me.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub status_AfterUpdate()
On Error GoTo Err_Handler

If Me.responsible = userData("Dept") And DCount("recordId", "tblPartTeam", "person = '" & Environ("username") & "' AND partNumber = '" & Me.partNumber & "'") > 0 Then GoTo goodToGo 'responsible and on CF team
If restrict(Environ("username"), TempVars!projectOwner) = False Then GoTo goodToGo 'is the user an owner
Me.ActiveControl = Me.ActiveControl.OldValue
Call snackBox("error", "Not today.", "Only Responsible person or PE can edit status", "frmPartDashboard")
Exit Sub

goodToGo:

If IsNull(Me.closeDate) = False And restrict(Environ("username"), TempVars!projectOwner, "Manager") Then
    Me.ActiveControl = Me.ActiveControl.OldValue
    Call snackBox("error", "You can't edit a closed step", "Only a project/service manager and edit a closed step - looks like that's not you, sorry.", "frmPartDashboard")
    Exit Sub
End If

If Me.ActiveControl.OldValue = "Closed" Then
    Me.closeDate = Null
    Call registerPartUpdates("tblPartSteps", Me.recordId, Me.closeDate.name, Me.closeDate.OldValue, Me.closeDate, Me.partNumber, Me.type, Me.partProjectId)
End If

Call registerPartUpdates("tblPartSteps", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.type, Me.partProjectId)
Me.lastUpdatedDate = Now()
Me.lastUpdatedBy = Environ("username")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub stepFiles_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmPartAttachments", , , "partStepId = " & Me.recordId
Form_frmPartAttachments.TpartNumber = Me.partNumber
Form_frmPartAttachments.TstepId = Me.recordId
Form_frmPartAttachments.TprojectId = Me.partProjectId
Form_frmPartAttachments.itemName.Caption = Me.stepType

'TEMPORARY RESTRICTION OVERRIDE
'Project engineers can add/delete any file while in beta testing
If restrict(Environ("username"), TempVars!projectOwner) = False Then Exit Sub 'is the bro an owner?

'FIRST: are you the right person for the job???
If Me.responsible = userData("Dept") And DCount("recordId", "tblPartTeam", "person = '" & Environ("username") & "' AND partNumber = '" & Me.partNumber & "'") > 0 Then Exit Sub 'if the bro is responsible AND CHECK IF ON CF TEAM
If restrict(Environ("username"), Nz(Me.responsible, TempVars!projectOwner), "Manager") = False Then Exit Sub  'is the bro a manager in the department of the "responsible" person?

Form_frmPartAttachments.remove.Visible = False
Form_frmPartAttachments.newAttachment.Visible = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub stepHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory", acNormal, , "[tableName] = 'tblPartSteps' AND [tableRecordId] = " & Me.recordId

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Public Sub sfrmPartDashLock(lockIt As Boolean)
On Error GoTo Err_Handler

Me.closeStep.Enabled = Not lockIt
Me.approvals.Enabled = Not lockIt
Me.nudgeApprovers.Enabled = Not lockIt
Me.nudgeResponsible.Enabled = Not lockIt

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
