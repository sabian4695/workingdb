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

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub closeStep_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False
If closeProjectStep(Me.recordId, "frmNMQDashboard") Then
    Call snackBox("success", "Success!", "Step Closed", "frmNMQDashboard")
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

Public Function setFilter()
On Error GoTo Err_Handler

Dim extFilt As String
If Me.showClosedToggle Then
    extFilt = "="
Else
    extFilt = "<>"
End If

Me.filter = "(recordId IN (Select tableRecordId FROM tblPartTrackingApprovals WHERE dept = '" & userData("Dept") & "' AND reqLevel = '" & userData("Level") & "' AND approvedOn is null) OR (responsible = '" & userData("Dept") & "')) AND " & _
    "[status] " & extFilt & " 'Closed'"
Me.FilterOn = True


'Who is the owner?
If IsNull(Me.partProjectId) Then Exit Function
Select Case DLookup("templateType", "tblPartProjectTemplate", "recordId = " & DLookup("projectTemplateId", "tblPartProject", "recordId = " & Nz(Me.partProjectId)))
    Case 1 'New Model
        TempVars.Add "projectOwner", "Project"
    Case 2 'Service
        TempVars.Add "projectOwner", "Service"
End Select

Exit Function
Err_Handler:
    Call handleError(Me.name, "setFilter", Err.DESCRIPTION, Err.number)
End Function

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.showClosedToggle = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
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

Private Sub showClosedToggle_Click()
On Error GoTo Err_Handler

setFilter

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Sub updateLastUpdate()
On Error GoTo Err_Handler

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
