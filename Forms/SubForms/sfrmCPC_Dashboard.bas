Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub approvals_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmCPC_Approvals", , , "[stepId] = " & Me.ID

Form_frmCPC_Approvals.newApproval.tag = Me.ID
Form_frmCPC_Approvals.projId = Me.projectId
Form_frmCPC_Approvals.itemName.Caption = Me.stepName
Form_frmCPC_Approvals.approvalNote.tag = Nz(Me.responsible)

Dim allowIt As Boolean
allowIt = Me.status <> "Closed"

Form_frmCPC_Approvals.approve.Enabled = allowIt
Form_frmCPC_Approvals.allowEdits = allowIt
Form_frmCPC_Approvals.sendFeedback.Enabled = allowIt
Form_frmCPC_Approvals.nudge.Enabled = allowIt
Form_frmCPC_Approvals.remove.Enabled = allowIt
Form_frmCPC_Approvals.newApproval.Enabled = allowIt
Form_frmCPC_Approvals.save.Enabled = allowIt

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnEditInfo_Click()
On Error GoTo Err_Handler

If restrict(Environ("username"), "CPC", "Supervisor", True) Then
    Call snackBox("error", "You can't edit a closed step", "Only a CPC manager can edit a step - looks like that's not you, sorry.", "frmCPC_Dashboard")
    Exit Sub
End If

'DoCmd.OpenForm "frmPartStepEdit", acNormal, , "ID = " & Me.ID

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub closeStep_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False
If closeCPCstep(Me.ID, "frmCPC_Dashboard") Then
    Call snackBox("success", "Success!", "Step Closed", "frmCPC_Dashboard")
    Me.Requery
    Me.refresh
    Call updateLastUpdate
End If

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

If IsNull(Me.closeDate) = False And restrict(Environ("username"), "CPC", "Manager") Then
    Me.dueDate = Me.dueDate.OldValue
    Call snackBox("error", "You can't edit a closed step", "Only a CPC manager and edit a closed step - looks like that's not you, sorry.", "frmCPC_Dashboard")
    Exit Sub
End If

Call registerCPCUpdates("tblCPC_Steps", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.projectId, Me.stepName)

Me.Requery
Call updateLastUpdate

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub newStep_Click()
On Error GoTo Err_Handler

If IsNull(Me.closeDate) = False And restrict(Environ("username"), Nz("CPC", ""), "Manager") Then
    Call snackBox("error", "You can't edit a closed gate", "Only a CPC manager and edit a closed gate - looks like that's not you, sorry.", "frmCPC_Dashboard")
    Exit Sub
End If

If MsgBox("By default, the added step will be placed AFTER the step you have selected currently. Shall I proceed?", vbYesNo, "Error") <> vbYes Then Exit Sub

DoCmd.OpenForm "frmCPC_NewStep"
Form_frmCPC_NewStep.projectId = Me.projectId
Form_frmCPC_NewStep.indexOrder = Me.indexOrder + 1

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub nudgeApprovers_Click()
On Error GoTo Err_Handler

Exit Sub

If restrict(Environ("username"), "CPC") = False Then GoTo theCorrectFellow 'is the bro an owner?
If Me.responsible = userData("Dept") And DCount("ID", "tblCPC_XFTeams", "memberName = '" & Environ("username") & "' AND projectId = " & Me.projectId) > 0 Then GoTo theCorrectFellow 'if the bro is responsible AND CHECK IF ON CF TEAM
If restrict(Environ("username"), Me.responsible, "Manager") = False Then GoTo theCorrectFellow  'is the bro a manager in the department of the "responsible" person?
MsgBox "Only the 'Responsible' person, their manager, or a CPC Engineer can nudge these approvers", vbCritical, "Woops"
Exit Sub
theCorrectFellow:

Dim db As Database
Set db = CurrentDb()
Dim rs3 As Recordset
Set rs3 = db.OpenRecordset("select * from tblCPC_Approvals where approvedOn is null and stepId = " & Me.ID)

If rs3.RecordCount = 0 Then
    MsgBox "No open approvals found!", vbInformation, "Hmm.."
    Exit Sub
End If

Dim stepTitle As String
stepTitle = Me.stepName

Dim sendTo As String
Dim body As String
body = emailContentGen("You've been nudged...", "Nudge Notification", "You've been nudged by " & getFullName() & " to approve this step", stepTitle, "Project Number: " & Me.projectId, "Requested By: " & rs3!requestedBy, "Requested On: " & CStr(rs3!requestedDate))

Do While Not rs3.EOF 'loop through all the approvals
    sendTo = Nz(rs3!approver)
    If sendTo = Environ("username") Then GoTo nextOne
    
    If Nz(sendTo) = "" Then 'if approver not specified, look through cross functional team and see if anyone matches
        Dim rs1 As Recordset, rs2 As Recordset
        Set rs1 = db.OpenRecordset("select * from tblCPC_XFTeams where projectId = " & Me.projectId)
        Do While Not rs1.EOF
            Set rs2 = db.OpenRecordset("select * from tblPermissions where user = '" & rs1!memberName & "'")
            If rs2!dept <> rs3!dept Or rs2!Level <> rs3!reqLevel Then GoTo nextOne 'dept/level dont match. skip this person
            sendTo = rs2!User
            If sendTo = Environ("username") Then
                MsgBox "can't nudge yourself", vbInformation, "Darn"
                GoTo nextOne
            End If
            If sendNotification(sendTo, 1, 2, "Please approve step " & Me.stepName, body, "CPC Project", Me.projectId) = True Then
                MsgBox "Notification sent to " & sendTo & "!", vbInformation, "Well done."
                Call registerCPCUpdates("tblCPC_Steps", Me.ID, "Nudge", "From: " & Environ("username"), "To: " & sendTo, Me.projectId, Me.stepName)
            End If
nextOne:
            rs1.MoveNext
        Loop
        If sendTo = "" Then MsgBox rs3!dept & " " & rs3!reqLevel & " not found on the cross functional team", vbInformation, "Sorry, can't do that"
    Else
        If sendNotification(sendTo, 1, 2, "Please approve step " & Me.stepName, body, "CPC Project", Me.projectId) = True Then
            MsgBox "Notification sent to " & sendTo & "!", vbInformation, "Well done."
            Call registerCPCUpdates("tblCPC_Steps", Me.ID, "Nudge", "From: " & Environ("username"), "To: " & sendTo, Me.projectId, Me.stepName)
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
    Call snackBox("error", "No can do", "Need a 'responsible' department to nudge", "frmCPC_Dashboard")
    Exit Sub
End If
notiSent = False

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset, rs2 As Recordset
Set rs1 = db.OpenRecordset("select * from tblCPC_XFTeams where projectId = " & Me.projectId)
Do While Not rs1.EOF
    Set rs2 = db.OpenRecordset("select * from tblPermissions where user = '" & rs1!memberName & "'")
    If rs2!dept <> Me.responsible Then GoTo nextOne 'if dept isn't the same, skip this user
    sendTo = rs2!User
    If sendTo = Environ("username") Then
        Call snackBox("error", "No can do", "You can't nudge yourself!", "frmCPC_Dashboard")
        GoTo nextOne
    End If

    stepTitle = Me.stepName

    Dim body As String
    body = emailContentGen("You've been nudged...", "Nudge Notification", "You've been nudged by " & getFullName() & " to complete this step", stepTitle, "Project Number: " & "FIND PROJ #", "Requested By: " & Environ("username"), "Requested On: " & CStr(Date))
    If sendNotification(sendTo, 1, 2, "Please complete step " & Me.stepName, body, "CPC Project", Me.projectId) = True Then
        Call snackBox("success", "Well done.", "Notification sent to " & sendTo & "!", "frmCPC_Dashboard")
        Call registerCPCUpdates("tblCPC_Steps", Me.ID, "Nudge", "From: " & Environ("username"), "To: " & sendTo, Me.projectId, Me.stepName)
        notiSent = True
    End If
nextOne:
    rs1.MoveNext
Loop

If Not notiSent Then Call snackBox("error", "Woops", "No one found", "frmCPC_Dashboard")

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

Dim extFilt As String
If Me.ActiveControl.Value Then
    extFilt = "="
    Me.lblDue.Caption = "Closed"
Else
    extFilt = "<>"
    Me.lblDue.Caption = "Due"
End If

Me.dueDate.Visible = Not Me.ActiveControl.Value
Me.filter = "[status] " & extFilt & " 'Closed'"
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
    Call registerCPCUpdates("tblCPC_Steps", Me.ID, "Status", "Not Started", "In Progress", Me.projectId, Me.stepName)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, "updateLastUpdate", Err.DESCRIPTION, Err.number)
End Sub

Private Sub status_AfterUpdate()
On Error GoTo Err_Handler

If Me.responsible = userData("Dept") And DCount("ID", "tblCPC_XFTeams", "memberName = '" & Environ("username") & "' AND projectId = " & Me.projectId) > 0 Then GoTo goodToGo 'responsible and on CF team
If restrict(Environ("username"), "CPC") = False Then GoTo goodToGo 'is the user an owner
Me.ActiveControl = Me.ActiveControl.OldValue
Call snackBox("error", "Not today.", "Only Responsible person or PE can edit status", "frmCPC_Dashboard")
Exit Sub

goodToGo:

If IsNull(Me.closeDate) = False And restrict(Environ("username"), "CPC", "Manager") Then
    Me.ActiveControl = Me.ActiveControl.OldValue
    Call snackBox("error", "You can't edit a closed step", "Only a CPC manager and edit a closed step - looks like that's not you, sorry.", "frmCPC_Dashboard")
    Exit Sub
End If

If Me.ActiveControl.OldValue = "Closed" Then
    Me.closeDate = Null
    Call registerCPCUpdates("tblCPC_Steps", Me.ID, Me.closeDate.name, Me.closeDate.OldValue, Me.closeDate, Me.projectId, Me.stepName)
End If

Call registerCPCUpdates("tblCPC_Steps", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.projectId, Me.stepName)
Me.lastUpdatedDate = Now()
Me.lastUpdatedBy = Environ("username")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub stepFiles_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmPartAttachments", , , "partStepId = " & Me.ID
Form_frmPartAttachments.TstepId = Me.ID
Form_frmPartAttachments.TprojectId = Me.projectId
Form_frmPartAttachments.itemName.Caption = Me.stepName

'TEMPORARY RESTRICTION OVERRIDE
'Project engineers can add/delete any file while in beta testing
If restrict(Environ("username"), "CPC") = False Then Exit Sub 'is the bro an owner?

'FIRST: are you the right person for the job???
If Me.responsible = userData("Dept") And DCount("ID", "tblCPC_XFTeams", "memberName = '" & Environ("username") & "' AND projectId = " & Me.projectId) > 0 Then Exit Sub 'if the bro is responsible AND CHECK IF ON CF TEAM
If restrict(Environ("username"), Nz(Me.responsible, "CPC"), "Manager") = False Then Exit Sub  'is the bro a manager in the department of the "responsible" person?

Form_frmPartAttachments.remove.Visible = False
Form_frmPartAttachments.newAttachment.Visible = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub stepHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory", acNormal, , "[tableName] = 'tblCPC_Steps' AND [stepId] = " & Me.ID

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

Private Sub stepNotes_AfterUpdate()
On Error GoTo Err_Handler

If IsNull(Me.closeDate) = False Then
    If restrict(Environ("username"), "CPC", "Manager") = True Then
        Me.ActiveControl = Me.ActiveControl.OldValue
        Call snackBox("error", "You can't edit a closed step", "Only a CPC manager and edit a closed step - looks like that's not you, sorry.", "frmCPC_Dashboard")
        Exit Sub
    End If
End If

Call registerCPCUpdates("tblCPC_Steps", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.projectId, Me.stepName)
Call updateLastUpdate

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
