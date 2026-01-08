Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub approvalHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory", acNormal, , "[tableName] = 'tblCPC_StepApprovals' AND [tableID] = " & Me.stepId

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub approvalNote_AfterUpdate()
On Error GoTo Err_Handler

Dim projId As Long
projId = DLookup("projectId", "tblCPC_Steps", "ID = " & Me.stepId)

Call registerCPCUpdates("tblCPC_StepApprovals", Me.stepId, "Approval", Me.ActiveControl.OldValue, Me.ActiveControl, Me.projId, Me.itemName.Caption)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub approvalNote_GotFocus()
On Error Resume Next
Me.approvalNote.Locked = restrict(Environ("username"), Me.dept, Me.reqLevel, True) 'let anyone above reqLevel approve
End Sub

Private Sub Detail_Paint()
On Error Resume Next

Dim itsNotMe As Boolean
itsNotMe = True

If Nz(Me.approvedOn) = "" And Not restrict(Environ("username"), Me.dept, Me.reqLevel, True) Then
    itsNotMe = False
End If

Me.approve.Transparent = itsNotMe
Me.sendFeedback.Transparent = itsNotMe
Me.nudge.Transparent = Nz(Me.approvedOn, False)

End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

If Not IsNull("CPC") Then
    If (restrict(Environ("username"), "CPC") = True) Then Me.newApproval.Visible = False
Else
    Me.newApproval.Visible = False
End If

Me.save.Visible = False
Me.dept.Locked = True
Me.reqLevel.Locked = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub approve_Click()
On Error GoTo Err_Handler

Dim rsStep As Recordset, db As Database
Dim errorText As String
errorText = ""

If Me.save.Visible Then errorText = "Please save first"
'is this the right person?
If restrict(Environ("username"), Nz(Me.dept), Nz(Me.reqLevel), True) Then errorText = "Only a " & Nz(Me.dept) & " " & Nz(Me.reqLevel) & " (or above) can approve this." 'let anyone above reqLevel approve

If errorText <> "" Then GoTo errorOut

'is there a file required? If so, only allow approvals after it is uploaded
Set db = CurrentDb()
Set rsStep = db.OpenRecordset("SELECT * FROM tblCPC_Steps WHERE ID = " & Me.stepId)

If Nz(rsStep!documentType, 0) <> 0 Then
    Dim rsAttach As Recordset, rsProjPNs As Recordset, rsUpdates As Recordset
    Set rsAttach = db.OpenRecordset("SELECT * FROM tblPartAttachmentsSP WHERE partstepId = " & rsStep!ID)
    
    'if 0 attachments then autofail
    If rsAttach.RecordCount = 0 Then
        errorText = "This step requires a file to be added to approve it"
        GoTo errorOut
    End If
    
    'if file found, check if it has been opened
    Set rsUpdates = db.OpenRecordset("SELECT * FROM tblCPC_UpdateTracking WHERE tableName = 'tblPartAttachmentsSP' AND columnName = 'Step Attachment' AND newData = 'Opened' AND tableID = " & rsAttach!ID)
    If rsUpdates.RecordCount = 0 Then
        If MsgBox("You have not opened the attached file. Are you sure you want to approve this step? Seems kinda lazy.", vbYesNo, "Check Yourself...") = vbNo Then
            errorText = "Approval cancelled because you are a good person"
            GoTo errorOut
        End If
    End If
End If

GoTo approveIt
errorOut:
Call snackBox("error", "Denied", errorText, Me.name)
Exit Sub


'---ALL CHECKS ARE OK, MOVE ON---
approveIt:
Me.approver = Environ("username")
Me.approvedOn = Now()

Dim stepType As String
stepType = Nz(Me.itemName.Caption)
Call registerCPCUpdates("tblCPC_StepApprovals", Me.stepId, "Approval", "", Me.approvedOn, Me.projId, Me.itemName.Caption)

Me.Dirty = False

'NOTIFY RESPONSIBLE PERSON
'---if fully approved only---
If getApprovalsCompleteCPC(Me.stepId) = getTotalApprovalsCPC(Me.stepId) Then
    Dim stepTitle As String, strSendTo As String
    stepTitle = Me.itemName.Caption

    Dim body As String, srtSendTo As String
    body = emailContentGen(stepTitle & " Step Approved", "WDB Step Fully Approved", "Step Approved: Ready to Close", stepTitle, "Last Approval: " & getFullName(), "", "")
    
    'try to find responsible person
    strSendTo = findDeptCPC(Me.projId, Me.approvalNote.tag)
    If strSendTo = "" Then 'if it can't find them, send to project owner (PE usually)
        strSendTo = findDeptCPC(Me.projId, "CPC")
    Else 'add project owner either way
        If Me.approvalNote.tag <> "CPC" Then strSendTo = strSendTo & "," & findDeptCPC(Me.projId, "CPC")
    End If
    
    If strSendTo <> "" Then Call sendNotification(strSendTo, 3, 2, stepTitle & " for CPC Project" & " Approved - Ready to Close", body, "CPC Project", Me.projId, True)
End If

Call snackBox("success", "Thank you!", "Step Approved", Me.name)

'clear variables
exit_handler:
On Error Resume Next
rsAttach.CLOSE: Set rsAttach = Nothing
rsProjPNs.CLOSE: Set rsProjPNs = Nothing
rsStep.CLOSE: Set rsStep = Nothing
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler

If Me.save.Visible Then
    Call snackBox("error", "Denied", "Please save first", Me.name)
    Cancel = True
    Exit Sub
End If

If CurrentProject.AllForms("sfrmCPC_Dashboard").IsLoaded Then forms.frmCPC_Dashboard.sfrmCPC_Dashboard.Form.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Unload", Err.DESCRIPTION, Err.number)
End Sub

Private Sub nudge_Click()
On Error GoTo Err_Handler

If Me.save.Visible Then
    Call snackBox("error", "Denied", "Please save first", Me.name)
    Exit Sub
End If

Dim sendTo As String, errorText As String
sendTo = Nz(Me.approver)
errorText = ""

If sendTo = Environ("username") Then errorText = "You can't nudge yourself"
If Nz(Me.approvedOn) = "" = False Then errorText = "This is already approved"
If errorText <> "" Then
    Call snackBox("error", "No can do!", errorText, Me.name)
    Exit Sub
End If

Dim db As Database
Set db = CurrentDb()

Dim stepTitle As String, body As String
stepTitle = Me.itemName.Caption
body = emailContentGen("You've been nudged...", "Nudge Notification", "You've been nudged by " & getFullName() & " to approve this step", stepTitle, "Requested By: " & Me.requestedBy, "Requested On: " & CStr(Me.requestedOn), "")

If Nz(sendTo) = "" Then 'a general Department / Level approval
    Dim rs1 As Recordset, rs2 As Recordset
    Set rs1 = db.OpenRecordset("select * from tblCPC_XFteams where projectId = " & Me.projId)
    Do While Not rs1.EOF
        Set rs2 = db.OpenRecordset("select * from tblPermissions where user = '" & rs1!memberName & "'")
        If rs2!dept <> Me.dept Or rs2!Level <> Me.reqLevel Then GoTo nextOne 'if dept/level isn't a match, this person isn't qualified.
        sendTo = rs2!User
        If sendTo = Environ("username") Then GoTo nextOne 'dont send a nudge to yourself
        If sendNotification(sendTo, 1, 2, "Please approve step " & stepTitle, body, "CPC Project", Me.projId) = True Then
            Call snackBox("success", "Let's get this done!", "Notification sent to " & sendTo & "!", Me.name)
            Call registerCPCUpdates("tblCPC_Steps", Me.stepId, "Nudge", "From: " & Environ("username"), "To: " & sendTo, Me.projId, Me.itemName.Caption)
        End If
nextOne:
        rs1.MoveNext
    Loop
    If sendTo = "" Then MsgBox Me.dept & " " & Me.reqLevel & " not found on the cross functional team", vbInformation, "Sorry, can't do that"
Else
    If sendNotification(sendTo, 1, 2, "Please approve step " & stepTitle, body, "CPC Project", Me.projId) = True Then
        Call snackBox("success", "Let's get this done!", "Notification sent to " & sendTo & "!", Me.name)
        Call registerCPCUpdates("tblCPC_Steps", Me.stepId, "Nudge", "From: " & Environ("username"), "To: " & sendTo, Me.projId, Me.itemName.Caption)
    End If
End If

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

Private Sub remove_Click()
On Error GoTo Err_Handler

If restrict(Environ("username"), "Project", "Supervisor", True) Then
    Call snackBox("error", "Denied", "Only Project Supervisors or Managers can edit this field", Me.name)
    Exit Sub
End If

Dim projId As Long
projId = DLookup("projectId", "tblCPC_Steps", "ID = " & Me.stepId)

If MsgBox("Are you sure you want to delete this?", vbYesNo, "Please confirm") = vbYes Then
    Call registerCPCUpdates("tblCPC_StepApprovals", Me.stepId, "Approval", Me.approvedOn, "Deleted", Me.projId, Me.itemName.Caption)
    dbExecute ("DELETE FROM tblCPC_StepApprovals WHERE [ID] = " & Me.ID)
    Me.Requery
    Me.save.Visible = False
    Me.dept.Locked = True
    Me.reqLevel.Locked = True
    Call snackBox("success", "Success!", "Approval deleted", Me.name)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub newApproval_Click()
On Error GoTo Err_Handler

Me.approvalHistory.SetFocus
Me.newApproval.Visible = False

Dim db As Database
Set db = CurrentDb()

Dim stepId
stepId = Me.newApproval.tag

db.Execute "INSERT INTO tblCPC_StepApprovals(stepId,requestedBy,requestedOn) VALUES (" & _
    stepId & ",'" & Environ("username") & "',#" & Now() & "#);"
TempVars.Add "approvalId", db.OpenRecordset("SELECT @@identity")(0).Value
Call registerCPCUpdates("tblCPC_StepApprovals", stepId, "Approval", "", "Created", Me.projId, Me.itemName.Caption)
Me.Requery
Me.filter = "ID = " & TempVars!approvalId
Me.FilterOn = True

'filter to nothing, goto new record, then unlock role dropdowns
Me.save.Visible = True
Me.dept.Locked = False
Me.reqLevel.Locked = False

Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Function validate() As Boolean

validate = False

Dim errorMsg As String
errorMsg = ""

If (Me.dept = "" Or IsNull(Me.dept)) Then errorMsg = "Department"
If (Me.reqLevel = "" Or IsNull(Me.reqLevel)) Then errorMsg = "Level"

If errorMsg <> "" Then
    Call snackBox("error", "Please fix", "Please fill out " & errorMsg, Me.name)
    Exit Function
End If

validate = True

End Function

Public Sub refresh_Click()
On Error GoTo Err_Handler

If Me.save.Visible Then
    Call snackBox("error", "Denied", "Please save first", Me.name)
    Exit Sub
End If

Me.Requery
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub save_Click()
On Error GoTo Err_Handler

If validate = False Then Exit Sub

Me.AllowAdditions = False
Me.filter = "[stepId] = " & Me.stepId
Me.FilterOn = True
Me.approvalHistory.SetFocus
Me.save.Visible = False
Me.dept.Locked = True
Me.reqLevel.Locked = True

Me.newApproval.Visible = True

Call snackBox("success", "Success!", "Approval saved", Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub sendFeedback_Click()
On Error GoTo Err_Handler

If Nz(Me.approvalNote) = "" Then
    Call snackBox("error", "Just a moment...", "Please enter a note to send feedback!", Me.name)
    Exit Sub
End If

Dim stepTitle As String, strSendTo As String, Notes As String
stepTitle = Me.itemName.Caption
Notes = Replace(Me.approvalNote, ",", ";")

If (restrict(Environ("username"), Me.dept, Me.reqLevel, True) = False) Then 'only let this approver send feedback
    Dim body As String, srtSendTo As String
    body = emailContentGen(stepTitle & " Step Feedback", "WDB Step Approval Rejection", "Feedback: " & Notes, stepTitle, "Sent by: " & getFullName(), "", "")
    
    strSendTo = findDeptCPC(Me.projId, Me.approvalNote.tag)
    If strSendTo = "" Then
        strSendTo = findDeptCPC(Me.projId, "CPC")
    Else
        If Me.approvalNote.tag <> "CPC" Then strSendTo = strSendTo & "," & findDeptCPC(Me.projId, "CPC")
    End If
    
    If Not IsNull(Me.approvedOn) Then
        Me.approver = Null
        Me.approvedOn = Null
        If Me.Dirty Then Me.Dirty = False
        Call registerCPCUpdates("tblCPC_StepApprovals", Me.stepId, "Approval", Me.approvedOn, "Approval Removed", Me.projId, Me.itemName.Caption)
    End If
    
    Call sendNotification(strSendTo, 7, 2, stepTitle & " Feedback", body, "CPC Project", Me.projId, True)
    Call registerCPCUpdates("tblCPC_StepApprovals", Me.stepId, "Approval", "Feedback Sent", "To: " & strSendTo & "; Notes: " & Notes, Me.projId, Me.itemName.Caption)
    Call snackBox("success", "Success!", "Feedback sent!", Me.name)
Else
    Call snackBox("error", "Denied", "Only the approver can send feedback", Me.name)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
