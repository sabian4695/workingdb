Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub approvalHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory", acNormal, , "[tableName] = 'tblPartTrackingApprovals' AND [tableRecordId] = " & Me.TtableRecordId

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub approvalNote_AfterUpdate()
On Error GoTo Err_Handler

Dim projId As Long
projId = DLookup("partProjectId", "tblPartSteps", "recordId = " & Me.TtableRecordId)

Call registerPartUpdates("tblPartTrackingApprovals", Me.TtableRecordId, "Approval", Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.itemName.Caption, projId)

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

If Not IsNull(TempVars!projectOwner) Then
    If (restrict(Environ("username"), TempVars!projectOwner) = True) Then Me.newApproval.Visible = False
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
Set rsStep = db.OpenRecordset("SELECT * FROM tblPartSteps WHERE recordId = " & Me.TtableRecordId)

If Nz(rsStep!documentType, 0) <> 0 Then
    Dim rsAttach As Recordset, rsAttStd As Recordset, rsProjPNs As Recordset, rsUpdates As Recordset
    Set rsAttach = db.OpenRecordset("SELECT * FROM tblPartAttachmentsSP WHERE partStepId = " & rsStep!recordId)
    Set rsAttStd = db.OpenRecordset("SELECT uniqueFile FROM tblPartAttachmentStandards WHERE recordId = " & rsStep!documentType)
    Set rsProjPNs = db.OpenRecordset("SELECT * from tblPartProjectPartNumbers WHERE projectId = " & rsStep!partProjectId)
    
    'if 0 attachments then autofail
    If rsAttach.RecordCount = 0 Then
        errorText = "This step requires a file to be added to approve it"
        GoTo errorOut
    End If
    
    'if file found, check if it has been opened
    Set rsUpdates = db.OpenRecordset("SELECT * FROM tblPartUpdateTracking WHERE tableName = 'tblPartAttachmentsSP' AND columnName = 'Step Attachment' AND newData = 'Opened' AND tableRecordId = " & rsAttach!ID)
    If rsUpdates.RecordCount = 0 Then
        If MsgBox("You have not opened the attached file. Are you sure you want to approve this step? Seems kinda lazy.", vbYesNo, "Check Yourself...") = vbNo Then
            errorText = "Approval cancelled because you are a good person"
            GoTo errorOut
        End If
    End If
    
    'If unique files are needed AND there is more than one part number, then check for an attachment for EACH part number
    If rsAttStd!uniqueFile And rsProjPNs.RecordCount > 0 Then
        'first, check primary PN
        rsAttach.FindFirst "partNumber = '" & rsStep!partNumber & "'"
        If rsAttach.noMatch Then
            errorText = "This step requires a file per related part number to be added to approve it. Nothing found for " & rsStep!partNumber
            GoTo errorOut
        End If
        
        'if file found, check if it has been opened
        Set rsUpdates = db.OpenRecordset("SELECT * FROM tblPartUpdateTracking WHERE tableName = 'tblPartAttachmentsSP' AND columnName = 'Step Attachment' AND newData = 'Opened' AND tableRecordId = " & rsAttach!ID)
        If rsUpdates.RecordCount = 0 Then
            If MsgBox("You have not opened the attached file for " & rsStep!partNumber & ". Are you sure you want to approve this step? Seems kinda lazy.", vbYesNo, "Check Yourself...") = vbNo Then
                errorText = "Approval cancelled because you are a good person"
                GoTo errorOut
            End If
        End If
        
        
        'then, check every childPartNumber
        Do While Not rsProjPNs.EOF
            rsAttach.FindFirst "partNumber = '" & rsProjPNs!childPartNumber & "'"
            If rsAttach.noMatch Then
                errorText = "This step requires a file per related part number to be added to approve it. Nothing found for " & rsProjPNs!childPartNumber
                GoTo errorOut
            End If
            
            'if file found, check if it has been opened
            Set rsUpdates = db.OpenRecordset("SELECT * FROM tblPartUpdateTracking WHERE tableName = 'tblPartAttachmentsSP' AND columnName = 'Step Attachment' AND newData = 'Opened' AND tableRecordId = " & rsAttach!ID)
            If rsUpdates.RecordCount = 0 Then
                If MsgBox("You have not opened the attached file for " & rsProjPNs!childPartNumber & ". Are you sure you want to approve this step? Seems kinda lazy.", vbYesNo, "Check Yourself...") = vbNo Then
                    errorText = "Approval cancelled because you are a good person"
                    GoTo errorOut
                End If
            End If
            
            rsProjPNs.MoveNext
        Loop
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

Dim stepType As String, projId As Long
projId = DLookup("partProjectId", "tblPartSteps", "recordId = " & Me.TtableRecordId)
stepType = Nz(Me.itemName.Caption)
Call registerPartUpdates("tblPartTrackingApprovals", Me.TtableRecordId, "Approval", "", Me.approvedOn, Me.partNumber, stepType, projId)
'Call notifyPE(Me.partNumber, "Approved", stepType)

Me.Dirty = False

'NOTIFY RESPONSIBLE PERSON
'---if fully approved only---
If getApprovalsComplete(Me.TtableRecordId, Me.partNumber) = getTotalApprovals(Me.TtableRecordId, Me.partNumber) Then
    Dim stepTitle As String, strSendTo As String
    stepTitle = Me.itemName.Caption

    Dim body As String, srtSendTo As String
    body = emailContentGen(stepTitle & " Step Approved", "WDB Step Fully Approved", "Step Approved: Ready to Close", stepTitle, "Part Number: " & Me.partNumber, "Last Approval: " & getFullName(), "", appName:="Part Project", appId:=Me.partNumber)

    'try to find responsible person
    strSendTo = findDept(Me.partNumber, Me.approvalNote.tag)
    If strSendTo = "" Then 'if it can't find them, send to project owner (PE usually)
        strSendTo = findDept(Me.partNumber, TempVars!projectOwner)
    Else 'add project owner either way
        If Me.approvalNote.tag <> TempVars!projectOwner Then strSendTo = strSendTo & "," & findDept(Me.partNumber, TempVars!projectOwner)
    End If

    Call sendNotification(strSendTo, 3, 2, stepTitle & " for " & Me.partNumber & " Approved - Ready to Close", body, "Part Project", Me.partNumber, True)
End If

Call snackBox("success", "Thank you!", "Step Approved", Me.name)

'clear variables
exit_handler:
On Error Resume Next
rsAttach.CLOSE: Set rsAttach = Nothing
rsAttStd.CLOSE: Set rsAttStd = Nothing
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

If CurrentProject.AllForms("sfrmPartDashboard").IsLoaded Then forms.frmPartDashboard.sfrmPartDashboard.Form.Requery

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
Dim projId As Long
projId = DLookup("partProjectId", "tblPartSteps", "recordId = " & Me.TtableRecordId)
stepTitle = Me.itemName.Caption
body = emailContentGen("You've been nudged...", "Nudge Notification", "You've been nudged by " & getFullName() & " to approve this step", stepTitle, "Part Number: " & Me.partNumber, "Requested By: " & Me.requestedBy, "Requested On: " & CStr(Me.requestedDate), appName:="Part Project", appId:=Me.partNumber)

If Nz(sendTo) = "" Then 'a general Department / Level approval
    Dim rs1 As Recordset, rs2 As Recordset
    Set rs1 = db.OpenRecordset("select * from tblPartTeam where partNumber = '" & Me.partNumber & "'")
    Do While Not rs1.EOF
        Set rs2 = db.OpenRecordset("select * from tblPermissions where user = '" & rs1!person & "'")
        If rs2!dept <> Me.dept Or rs2!Level <> Me.reqLevel Then GoTo nextOne 'if dept/level isn't a match, this person isn't qualified.
        sendTo = rs2!User
        If sendTo = Environ("username") Then GoTo nextOne 'dont send a nudge to yourself
        If sendNotification(sendTo, 1, 2, "Please approve step " & stepTitle & " for " & Me.partNumber, body, "Part Project", Me.partNumber) = True Then
            Call snackBox("success", "Let's get this done!", "Notification sent to " & sendTo & "!", Me.name)
            Call registerPartUpdates("tblPartSteps", Me.TtableRecordId, "Nudge", "From: " & Environ("username"), "To: " & sendTo, Me.partNumber, stepTitle, projId)
        End If
nextOne:
        rs1.MoveNext
    Loop
    If sendTo = "" Then MsgBox Me.dept & " " & Me.reqLevel & " not found on the cross functional team", vbInformation, "Sorry, can't do that"
Else
    If sendNotification(sendTo, 1, 2, "Please approve step " & stepTitle & " for " & Me.partNumber, body, "Part Project", Me.partNumber) = True Then
        Call snackBox("success", "Let's get this done!", "Notification sent to " & sendTo & "!", Me.name)
        Call registerPartUpdates("tblPartSteps", Me.TtableRecordId, "Nudge", "From: " & Environ("username"), "To: " & sendTo, Me.partNumber, stepTitle, projId)
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

If restrict(Environ("username"), TempVars!projectOwner, "Supervisor", True) Then
    Call snackBox("error", "Denied", "Only project/service Supervisors or Managers can edit this field", Me.name)
    Exit Sub
End If

Dim projId As Long
projId = DLookup("partProjectId", "tblPartSteps", "recordId = " & Me.TtableRecordId)

If MsgBox("Are you sure you want to delete this?", vbYesNo, "Please confirm") = vbYes Then
    Call registerPartUpdates("tblPartTrackingApprovals", Me.TtableRecordId, "Approval", Me.approvedOn, "Deleted", Me.partNumber, Me.itemName.Caption, projId)
    dbExecute ("DELETE FROM tblPartTrackingApprovals WHERE [recordId] = " & Me.recordId)
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

Dim projId As Long
projId = DLookup("partProjectId", "tblPartSteps", "recordId = " & Me.TtableRecordId)

db.Execute "INSERT INTO tblPartTrackingApprovals(partNumber,tableName,tableRecordId,requestedBy,requestedDate) VALUES ('" & _
    Me.TpartNumber & "','" & Me.TtableName & "'," & Me.TtableRecordId & ",'" & Environ("username") & "',#" & Now() & "#);"
TempVars.Add "approvalId", db.OpenRecordset("SELECT @@identity")(0).Value
Call registerPartUpdates("tblPartTrackingApprovals", Me.TtableRecordId, "Approval", "", "Created", Me.TpartNumber, Me.itemName.Caption, projId)
Me.Requery
Me.filter = "recordId = " & TempVars!approvalId
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
Me.filter = "[tableRecordId] = " & Me.tableRecordId & " AND [tableName] = 'tblPartSteps'"
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

Dim projId As Long
projId = DLookup("partProjectId", "tblPartSteps", "recordId = " & Me.TtableRecordId)

Dim stepTitle As String, strSendTo As String, Notes As String
stepTitle = Me.itemName.Caption
Notes = Replace(Me.approvalNote, ",", ";")

If (restrict(Environ("username"), Me.dept, Me.reqLevel, True) = False) Then 'only let this approver send feedback
    Dim body As String, srtSendTo As String
    body = emailContentGen(stepTitle & " Step Feedback", "WDB Step Approval Rejection", "Feedback: " & Notes, stepTitle, "Part Number: " & Me.partNumber, "Sent by: " & getFullName(), "", appName:="Part Project", appId:=Me.partNumber)
    
    strSendTo = findDept(Me.partNumber, Me.approvalNote.tag)
    If strSendTo = "" Then
        strSendTo = findDept(Me.partNumber, TempVars!projectOwner)
    Else
        If Me.approvalNote.tag <> TempVars!projectOwner Then strSendTo = strSendTo & "," & findDept(Me.partNumber, TempVars!projectOwner)
    End If
    
    If Not IsNull(Me.approvedOn) Then
        Me.approver = Null
        Me.approvedOn = Null
        If Me.Dirty Then Me.Dirty = False
        Call registerPartUpdates("tblPartTrackingApprovals", Me.TtableRecordId, "Approval", Me.approvedOn, "Approval Removed", Me.partNumber, stepTitle, projId)
    End If
    
    Call sendNotification(strSendTo, 7, 2, stepTitle & " for " & Me.partNumber & " Feedback", body, "Part Project", Me.partNumber, True)
    Call registerPartUpdates("tblPartTrackingApprovals", Me.TtableRecordId, "Approval", "Feedback Sent", "To: " & strSendTo & "; Notes: " & Notes, Me.partNumber, stepTitle, projId)
    Call snackBox("success", "Success!", "Feedback sent!", Me.name)
Else
    Call snackBox("error", "Denied", "Only the approver can send feedback", Me.name)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
