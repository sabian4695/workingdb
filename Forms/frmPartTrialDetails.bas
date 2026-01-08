Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub alternateMaterial_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTrials", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name, Nz(Me.projectId))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub creator_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTrials", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name, Nz(Me.projectId))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub cycleTime_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTrials", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name, Nz(Me.projectId))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub failureMode_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTrials", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name, Nz(Me.projectId))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub failureType_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTrials", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name, Nz(Me.projectId))

Me.failureMode.RowSource = "SELECT recordid, trialFailureTypeMode, trialFailureType From tblDropDownsSP WHERE trialFailureType Is Not Null AND trialFailureTypeMode = " & Me.failureType

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

If Not Me.Dirty Then Exit Sub
If Not restrict(Environ("username"), "Processing") Or Not restrict(Environ("username"), "Project") Then Exit Sub 'if processing engineer, then OK

MsgBox "You must be Processing or Project to edit", vbCritical, "Nope"
Me.Undo

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_BeforeUpdate", Err.DESCRIPTION, Err.number)
End Sub

Function validate()
On Error GoTo Err_Handler

validate = False

If IsNull(Me.recordId) Then
    validate = True
    Exit Function
End If

Dim errorArray As Collection
Set errorArray = New Collection

If IsNull(Me.trialDate) Then errorArray.Add "Trial Date"
If IsNull(Me.press) Then errorArray.Add "Press"

If Me.trialResult = 2 Then 'NG
    Me.lblFailureMode.Visible = True
    Me.lblFailureType.Visible = True
    Me.failureMode.Visible = True
    Me.failureType.Visible = True
    Me.rcFailure.Visible = True
    
    If Nz(Me.failureType, "") = "" Then errorArray.Add "Failure Type"
    If Nz(Me.failureMode, "") = "" Then errorArray.Add "Failure Mode"
End If

If errorArray.count > 0 Then
    Dim errorTxtLines As String, element
    errorTxtLines = ""
    For Each element In errorArray
        errorTxtLines = errorTxtLines & vbNewLine & element
    Next element
    
    MsgBox "Please fill out these items: " & vbNewLine & errorTxtLines, vbOKOnly, "No can do!"
    Exit Function
End If

validate = True

Exit Function
Err_Handler:
    Call handleError(Me.name, "validate", Err.DESCRIPTION, Err.number)
End Function

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim processE As Boolean, projectE As Boolean
processE = Not restrict(Environ("username"), "Processing")
projectE = Not restrict(Environ("username"), "Project")

'PE and Proc E can create/edit. ONLY Proc E can edit Status
Me.remove.Enabled = processE Or projectE

Me.Ttrial.Locked = Not (processE Or projectE)
Me.trialDate.Locked = Not (processE Or projectE)
Me.press.Locked = Not (processE Or projectE)
Me.trialStatus.Locked = Not processE
Me.trialResult.Locked = Not (processE Or projectE)
Me.trialReason.Locked = Not (processE Or projectE)
Me.alternateMaterial.Locked = Not (processE Or projectE)

Me.processingEngineer.Locked = Not processE
Me.oracleJob.Locked = Not processE
Me.cycleTime.Locked = Not processE
Me.runnerWeight.Locked = Not processE
Me.pieceWeight.Locked = Not processE
Me.failureMode.Locked = Not processE
Me.failureType.Locked = Not processE
Me.Notes.Locked = Not processE

If Me.trialResult = 2 Then 'NG
    Me.lblFailureMode.Visible = True
    Me.lblFailureType.Visible = True
    Me.failureMode.Visible = True
    Me.failureType.Visible = True
    Me.rcFailure.Visible = True
    If Nz(Me.failureType, "") <> "" Then Me.failureMode.RowSource = "SELECT recordid, trialFailureTypeMode, trialFailureType From tblDropDownsSP WHERE trialFailureType Is Not Null AND trialFailureTypeMode = " & Me.failureType
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False
If validate = False Then
    Cancel = True
    Exit Sub
End If


Form_frmPartTrialTracker.Requery
Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Unload", Err.DESCRIPTION, Err.number)
End Sub

Private Sub notes_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTrials", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name, Nz(Me.projectId))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openDetails_Click()
On Error GoTo Err_Handler

openPartProject (Me.partNumber)
DoCmd.CLOSE acForm, "frmPartTrialTracker"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openMasterSetup_Click()
On Error GoTo Err_Handler

Form_DASHBOARD.partNumberSearch = Me.partNumber
Form_DASHBOARD.cnlMasterSetups.SetFocus
Form_DASHBOARD.openMasterSetups

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub oracleJob_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTrials", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name, Nz(Me.projectId))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub partIssues_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmPartIssues", , , "partNumber = '" & Me.partNumber & "' AND [closeDate] is null"
Form_frmPartIssues.fltPartNumber = Me.partNumber

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub pieceWeight_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTrials", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name, Nz(Me.projectId))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub press_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTrials", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name, Nz(Me.projectId))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub processingEngineer_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTrials", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name, Nz(Me.projectId))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub remove_Click()
On Error GoTo Err_Handler

If (restrict(Environ("username"), "Processing") = True) And (restrict(Environ("username"), "Project") = True) Then
    MsgBox "Only Processing can remove this", vbCritical, "Denied"
    Exit Sub
End If

If MsgBox("Are you sure you want to delete this trial?", vbYesNo, "Please confirm") = vbYes Then
    Call registerPartUpdates("tblPartTrials", Me.recordId, "Trial", Me.trialDate, "Deleted", Me.partNumber, Me.name, Nz(Me.projectId))
    dbExecute ("DELETE FROM tblPartTrials WHERE [recordId] = " & Me.recordId)
    DoCmd.CLOSE acForm, "frmPartTrialDetails"
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub runnerWeight_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTrials", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name, Nz(Me.projectId))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub save_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False
If validate = False Then Exit Sub

DoCmd.CLOSE acForm, "frmPartTrialDetails"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub searchPN_Click()
On Error GoTo Err_Handler

Form_DASHBOARD.partNumberSearch = Me.partNumber
Form_DASHBOARD.filterbyPN_Click
Form_DASHBOARD.SetFocus

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub testFiles_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmPartAttachments", , , "partNumber = '" & Me.partNumber & "' AND docTypeId = 32"
Form_frmPartAttachments.TpartNumber = Me.partNumber
Form_frmPartAttachments.TtestId = Me.recordId
Form_frmPartAttachments.TprojectId = Nz(Me.projectId, 0)
Form_frmPartAttachments.itemName.Caption = "Part Trial"
Form_frmPartAttachments.secondaryType = 32

Form_frmPartAttachments.newAttachment.Visible = (DCount("recordId", "tblPartAttachmentsSP", "partNumber = '" & Me.partNumber & "' AND documentType = 32") = 0)

Form_frmPartAttachments.newAttachment.Enabled = Not restrict(Environ("username"), "Processing")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub trialHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory", acNormal, , "[tableName] = 'tblPartTrials' AND [partNumber] = '" & Me.partNumber & "' AND [tableRecordId] = " & Me.recordId

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub trialDate_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTrials", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name, Nz(Me.projectId))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub trialNotes_Click()
On Error GoTo Err_Handler

openPath (mainFolder(Me.ActiveControl.name) & "?FilterField1=Part_x0020_Number&FilterValue1=" & Me.partNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub trialReason_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTrials", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name, Nz(Me.projectId))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub trialResult_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTrials", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name, Nz(Me.projectId))

If Me.trialResult = 2 Then 'NG
    Me.lblFailureMode.Visible = True
    Me.lblFailureType.Visible = True
    Me.failureMode.Visible = True
    Me.failureType.Visible = True
    Me.rcFailure.Visible = True
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub trialStatus_AfterUpdate()
On Error GoTo Err_Handler

If Me.trialStatus = 3 Then 'complete
    'in order to mark a trial as complete, you must have cycle time / runner w / piece w - and you must add a "trial Result"
    
    Dim errorArray As Collection
    Set errorArray = New Collection
    
    If Nz(Me.cycleTime, 0) = 0 Then errorArray.Add "Cycle Time"
    If Nz(Me.pieceWeight, 0) = 0 Then errorArray.Add "Piece Weight"
    If Nz(Me.trialResult, 0) = 0 Then errorArray.Add "Trial Result"
    If Nz(Me.Notes, "") = "" Then errorArray.Add "Notes"
    If IsNull(Me.runnerWeight) Then errorArray.Add "Runner Weight"
    If IsNull(Me.trialReason) Then errorArray.Add "Trial Reason"
    
    If errorArray.count > 0 Then
        Dim errorTxtLines As String, element
        errorTxtLines = ""
        For Each element In errorArray
            errorTxtLines = errorTxtLines & vbNewLine & element
        Next element
        
        MsgBox "Please fill out these items: " & vbNewLine & errorTxtLines, vbOKOnly, "Fix this first."
        
        GoTo doNotComplete
    End If
    
    If validate = False Then GoTo doNotComplete
    
    'if PPH is less than quoted, ask for reason
    If Nz(Me.currentSPH, 0) < Nz(Me.shotsPerHour, 0) Then
        Dim redReason
        
enterStuff:
        Dim X
        X = InputBox("Please enter a reason for SPH being less previous", "Please add information")
        If StrPtr(X) = 0 Then
            Me.trialStatus = Me.trialStatus.OldValue
            Exit Sub
        End If
        If X = "" Then GoTo enterStuff
        
        redReason = X
        
        Call registerPartUpdates("tblPartTrials", Me.recordId, "SPH Decrease", Nz(Me.shotsPerHour, 0) & " -> " & Round(Nz(Me.currentSPH, 0), 2), "REASON: " & redReason, Me.partNumber, Me.name, Nz(Me.projectId))
    End If
    
    'check VALIDATE function, so that failure modes are input if needed
    If Me.trialResult = 2 Then 'then send email to team if trial failed
        Dim emailBody As String, subjectLine As String
        subjectLine = Me.toolNumber & " Trial Failure"
        emailBody = generateHTML(subjectLine, Me.trialReason.column(1) & " trial for " & partNumber & " resulted NG. </br>" & _
            "Notes: " & Me.Notes, "Trial Failure", "Trial Date: " & Me.trialDate, "Failure Type: " & Me.failureType.column(1), "Failure Mode: " & Me.failureMode.column(2), appName:="Trial", appId:=Me.recordId)
        
        Call sendNotification(grabPartTeam(partNumber, onlyEngineers:=True), 9, 3, partNumber & " Trial Failure", emailBody, "Trial", Me.recordId, True)
    End If
    
End If

Call registerPartUpdates("tblPartTrials", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name, Nz(Me.projectId))

Exit Sub
doNotComplete:
Me.trialStatus = Me.trialStatus.OldValue

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
