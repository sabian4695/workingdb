Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub allHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory", acNormal, , "[partNumber] = '" & Me.partNumber & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub approvalReport_Click()
On Error GoTo Err_Handler

DoCmd.OpenReport "rptPartApprovals", acViewPreview, , "partProjectId = " & Me.recordId

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub assPartNumber_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmPartProjectPartNumbers", acNormal, , "[projectId] = " & Me.recordId

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub docHisSearch_Click()
On Error GoTo Err_Handler

Form_DASHBOARD.partNumberSearch = Me.partNumber.Value
Call openDocumentHistoryFolder(Me.partNumber.Value)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnExportAIF_Click()
On Error GoTo Err_Handler

If DCount("recordId", "tblPartProjectPartNumbers", "projectId = " & Me.recordId) > 0 Then
    TempVars.Add "partDashAction", "AIF"
    DoCmd.OpenForm "frmPartProjectPartNumbers", acNormal, , "[projectId] = " & Me.recordId
    Form_frmPartProjectPartNumbers.doAction.Visible = True
    Form_frmPartProjectPartNumbers.doActionMaster.Visible = True
Else
    Call autoUploadAIF(Me.partNumber)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub findPOs_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmPOsearch", , , "capNum = '" & Me.projectCapitalNumber & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler
If CurrentProject.AllForms("frmPartInformation").IsLoaded Then DoCmd.CLOSE acForm, "frmPartInformation"
If CurrentProject.AllForms("frmPartIssues").IsLoaded Then DoCmd.CLOSE acForm, "frmPartIssues"
If CurrentProject.AllForms("frmPartMeetings").IsLoaded Then DoCmd.CLOSE acForm, "frmPartMeetings"
If CurrentProject.AllForms("frmPartTestingTracker").IsLoaded Then DoCmd.CLOSE acForm, "frmPartTesting"
If CurrentProject.AllForms("frmPartAttachments").IsLoaded Then DoCmd.CLOSE acForm, "frmPartAttachments"
If CurrentProject.AllForms("frmPartChangeTemplateSelector").IsLoaded Then DoCmd.CLOSE acForm, "frmPartChangeTemplateSelector"
If CurrentProject.AllForms("frmNewPartStep").IsLoaded Then DoCmd.CLOSE acForm, "frmNewPartStep"
If CurrentProject.AllForms("frmPartStepEdit").IsLoaded Then DoCmd.CLOSE acForm, "frmPartStepEdit"
If CurrentProject.AllForms("frmPartTrackingApprovals").IsLoaded Then DoCmd.CLOSE acForm, "frmPartTrackingApprovals"
If CurrentProject.AllForms("frmPartProjectPartNumbers").IsLoaded Then DoCmd.CLOSE acForm, "frmPartProjectPartNumbers"

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Unload", Err.DESCRIPTION, Err.number)
End Sub

Private Sub masterSchedule_Click()
On Error GoTo Err_Handler

Dim db As Database
Dim rs As Recordset

Dim errorbool As Boolean
errorbool = False

Set db = CurrentDb()
Set rs = db.OpenRecordset("SELECT * FROM tblPartAttachmentsSP WHERE partProjectId = " & Me.recordId & " AND documentType = 2")

If rs.RecordCount = 0 Then errorbool = True

rs.MoveLast
If rs!fileStatus <> "Uploaded" Then errorbool = True

If errorbool Then
    MsgBox "Master Setup not yet uploaded", vbInformation, "Hmmm, that's weird...?"
    GoTo exitFunction
End If

Application.FollowHyperlink rs!directLink

If Nz(rs!testId, 0) = 0 Then
    Call registerPartUpdates("tblPartAttachmentsSP", rs!ID, "Step Attachment", rs!attachName, "Opened", Me.partNumber, "Opened File", Me.recordId)
Else
    Call registerPartUpdates("tblPartAttachmentsSP", rs!ID, "Test Attachment", rs!attachName, "Opened", Me.partNumber, "Opened File", Me.recordId)
End If

exitFunction:

rs.CLOSE
Set rs = Nothing
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openAssy_Click()
On Error GoTo Err_Handler

Call openPartProject(Right(Me.ActiveControl.Caption, 5))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openItemReport_Click()
On Error GoTo Err_Handler

'if related parts found on this project, check what part number to open
Form_DASHBOARD.partNumberSearch = Me.partNumber

If DCount("recordId", "tblPartProjectPartNumbers", "projectId = " & Me.recordId) > 0 Then
    TempVars.Add "partDashAction", "rptPartOpenIssues"
    DoCmd.OpenForm "frmPartProjectPartNumbers", acNormal, , "[projectId] = " & Me.recordId
    Form_frmPartProjectPartNumbers.doAction.Visible = True
    Form_frmPartProjectPartNumbers.doActionMaster.Visible = True
Else
    DoCmd.OpenReport "rptPartOpenIssues", acViewPreview, , "partNumber = '" & Me.partNumber & "'"
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openPlannerECO_Click()
On Error GoTo Err_Handler

If Nz(Me.projectPlannerECO) = "" Then
    Call snackBox("error", "No ECO", "No ECO found!!", "frmPartDashboard")
    Exit Sub
End If

DoCmd.OpenForm "frmECOs", , , "[CHANGE_NOTICE] = '" & UCase(Me.projectPlannerECO) & "'"
Form_frmECOs.ECOsrch = Me.projectPlannerECO

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openProgram_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmProgramReview"

Dim Program
Program = Me.modelCode

Form_frmProgramReview.txtFilterInput.Value = Me.modelCode
Form_frmProgramReview.filterByProgram_Click

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openTransferECO_Click()
On Error GoTo Err_Handler

If Nz(Me.projectTransferECO) = "" Then
    Call snackBox("error", "No ECO", "No ECO found!", "frmPartDashboard")
    Exit Sub
End If

DoCmd.OpenForm "frmECOs", , , "[CHANGE_NOTICE] = '" & UCase(Me.projectTransferECO) & "'"
Form_frmECOs.ECOsrch = Me.projectTransferECO

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub partInfoReport_Click()
On Error GoTo Err_Handler

'if related parts found on this project, check what part number to open
Form_DASHBOARD.partNumberSearch = Me.partNumber

If DCount("recordId", "tblPartProjectPartNumbers", "projectId = " & Me.recordId) > 0 Then
    TempVars.Add "partDashAction", "rptPartInformation"
    DoCmd.OpenForm "frmPartProjectPartNumbers", acNormal, , "[projectId] = " & Me.recordId
    Form_frmPartProjectPartNumbers.doAction.Visible = True
    Form_frmPartProjectPartNumbers.doActionMaster.Visible = True
Else
    DoCmd.OpenReport "rptPartInformation", acViewPreview, , "[partNumber]='" & Me.partNumber & "'"
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub partTrackingHelp_Click()
On Error GoTo Err_Handler

Dim dept As String
dept = userData("Dept")

Select Case dept
    Case "Service", "Project"
        Call openPath(mainFolder(Me.ActiveControl.name))
    Case "Tooling"
        Call openPath(mainFolder(Me.ActiveControl.name & "_Tooling"))
    Case Else
        Call openPath(mainFolder(Me.ActiveControl.name & "_Universal"))
End Select

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub projectPlannerECO_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartProject", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.recordId)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub projectTabs_Change()
On Error GoTo Err_Handler

Select Case Me.projectTabs.Value
    Case 2 'open issues
        'if related parts found on this project, check what part number to open
        Form_DASHBOARD.partNumberSearch = Me.partNumber
        If DCount("recordId", "tblPartProjectPartNumbers", "projectId = " & Me.recordId) > 0 Then
            TempVars.Add "partDashAction", "frmPartIssues"
            DoCmd.OpenForm "frmPartProjectPartNumbers", acNormal, , "[projectId] = " & Me.recordId
            Form_frmPartProjectPartNumbers.doAction.Visible = True
            Form_frmPartProjectPartNumbers.doActionMaster.Visible = True
        Else
            If Me.frmPartIssues.SourceObject = "" Then
                Me.frmPartIssues.SourceObject = "frmPartIssues"
                Me.frmPartIssues.LinkChildFields = ""
                Me.frmPartIssues.LinkMasterFields = ""
            End If
            Form_frmPartIssues.Form.filter = "partNumber = '" & Me.partNumber & "' AND [closeDate] is null"
            Form_frmPartIssues.Form.FilterOn = True
            Form_frmPartIssues.Form.Controls("fltPartNumber") = Me.partNumber
            Me.frmPartIssues.Visible = True
        End If
    Case 3 'meetings
        If Me.frmPartMeetings.SourceObject = "" Then
            Me.frmPartMeetings.SourceObject = "frmPartMeetings"
            Me.frmPartMeetings.LinkChildFields = ""
            Me.frmPartMeetings.LinkMasterFields = ""
        End If
        Me.frmPartMeetings.Form.filter = "partNum = '" & Me.partNumber & "'"
        Me.frmPartMeetings.Form.FilterOn = True
        Form_frmPartMeetings.TpartNumber = Me.partNumber
        Form_frmPartMeetings.TprojectId = Form_frmPartDashboard.recordId
        Me.frmPartMeetings.Visible = True
    Case 4 'testing
        If Me.frmPartTestingTracker.SourceObject = "" Then
            Me.frmPartTestingTracker.SourceObject = "frmPartTestingTracker"
            Me.frmPartTestingTracker.LinkChildFields = ""
            Me.frmPartTestingTracker.LinkMasterFields = ""
        End If
        Me.frmPartTestingTracker.Form.filter = "partNumber = '" & Me.partNumber & "'"
        Me.frmPartTestingTracker.Form.FilterOn = True
        Form_frmPartTestingTracker.fltPartNumber = Me.partNumber
        Me.frmPartTestingTracker.Visible = True
    Case 5 'trials
        If Me.frmPartTrialTracker.SourceObject = "" Then
            Me.frmPartTrialTracker.SourceObject = "frmPartTrialTracker"
            Me.frmPartTrialTracker.LinkChildFields = ""
            Me.frmPartTrialTracker.LinkMasterFields = ""
        End If
        Me.frmPartTrialTracker.Form.filter = "partNumber = '" & Me.partNumber & "'"
        Me.frmPartTrialTracker.Form.FilterOn = True
        Form_frmPartTrialTracker.fltPartNumber = Me.partNumber
        Me.frmPartTrialTracker.Visible = True
    Case 6 'automation
        If Me.frmPartAutomationGates.SourceObject = "" Then
            Me.frmPartAutomationGates.SourceObject = "frmPartAutomationGates"
            Me.frmPartAutomationGates.LinkChildFields = ""
            Me.frmPartAutomationGates.LinkMasterFields = ""
        End If
        Me.frmPartAutomationGates.Form.filter = "partNumber = '" & Me.partNumber & "'"
        Me.frmPartAutomationGates.Form.FilterOn = True
        Form_frmPartAutomationGates.fltPartNumber = Me.partNumber
        Me.frmPartAutomationGates.Visible = True
    Case 7 'packaging
        If Me.frmPackagingDetails.SourceObject = "" Then
            Me.frmPackagingDetails.SourceObject = "frmPackagingDetails"
            Me.frmPackagingDetails.LinkChildFields = ""
            Me.frmPackagingDetails.LinkMasterFields = ""
        End If
        Me.frmPackagingDetails.Form.filter = "partNumber = '" & Me.partNumber & "'"
        Me.frmPackagingDetails.Form.FilterOn = True
        Me.frmPackagingDetails.Visible = True
    Case 8 'attachments
        If Me.frmPartAttachments.SourceObject = "" Then
            Me.frmPartAttachments.SourceObject = "frmPartAttachments"
            Me.frmPartAttachments.LinkChildFields = ""
            Me.frmPartAttachments.LinkMasterFields = ""
        End If
        Me.frmPartAttachments.Form.filter = "partProjectId = " & Me.recordId
        Me.frmPartAttachments.Form.FilterOn = True
        Form_frmPartAttachments.TpartNumber = Me.partNumber
        Form_frmPartAttachments.TprojectId = Me.recordId
        Form_frmPartAttachments.newAttachment.Visible = False
        Me.frmPartAttachments.Visible = True
End Select

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub projectTransferECO_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartProject", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.recordId)

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

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim partNum
partNum = TempVars!partNumber

Dim db As Database
Set db = CurrentDb()
Dim rsProjects As Recordset, projId
Set rsProjects = db.OpenRecordset("SELECT * from tblPartProject WHERE partNumber = '" & partNum & "'")

If rsProjects.RecordCount = 0 Then 'look for related projects
    projId = Nz(DLookup("projectId", "tblPartProjectPartNumbers", "childPartNumber = '" & partNum & "'"))
    If projId <> "" Then Set rsProjects = db.OpenRecordset("SELECT * from tblPartProject WHERE recordId = " & projId)
End If

If rsProjects.RecordCount = 0 Then
initialize:
    If restrict(Environ("username"), "Project") And restrict(Environ("username"), "Service") Then
        MsgBox "Only project/service engineers can set up a new part project", vbCritical, "No project created yet"
        GoTo killDash
    End If
    If MsgBox("Are you sure you want to set up a new part project for " & partNum & "?", vbYesNo, "Nothing found") = vbYes Then DoCmd.OpenForm "frmPartInitialize"
    GoTo killDash
End If

'Grab the latest project, check if it's open.
'if it's open, open it!! if not, check if you want to initialize new project or open the closed project

rsProjects.MoveLast
Dim recordId
recordId = rsProjects("recordId")
partNum = rsProjects("partNumber")

Dim totalSteps, closeSteps
totalSteps = DCount("recordId", "tblPartSteps", "partProjectId = " & recordId)
closeSteps = DCount("closeDate", "tblPartSteps", "partProjectId = " & recordId)

If closeSteps = totalSteps Then If MsgBox("The most recent project for this part is closed. Click YES to view the closed project, click NO to initialize a new one.", vbYesNo, "Just let me know") = vbNo Then GoTo initialize

Me.filter = "recordId = " & recordId
Me.FilterOn = True

Dim gateId As Long 'show steps for current open gate
gateId = Nz(DMin("[partGateId]", "tblPartSteps", "partProjectId = " & Me.recordId & " AND [status] <> 'Closed'"), DMin("[partGateId]", "tblPartSteps", "partProjectId = " & Me.recordId))

Form_sfrmPartDashboard.RecordSource = _
"SELECT tblPartSteps.*, dueDay([dueDate],[closeDate]) AS [Due Day], tblPartStepActions.stepAction " & _
"FROM tblPartSteps LEFT JOIN tblPartStepActions ON tblPartSteps.stepActionId = tblPartStepActions.recordId " & _
"WHERE tblPartSteps.partProjectId = " & recordId

Me.sfrmPartDashboard.Form.filter = "partGateId = " & gateId & " " & "AND [status] <> 'Closed'"
Me.sfrmPartDashboard.Form.FilterOn = True
Me.sfrmPartDashboard.Form.OrderBy = "indexOrder"
Me.sfrmPartDashboard.Form.OrderByOn = True

With Me.sfrmPartDashboardDates.Form.RecordsetClone 'set focus on current open gate on gate dates form
    .FindFirst "recordId = " & gateId
    Me.sfrmPartDashboardDates.Form.Bookmark = .Bookmark
End With
Call Form_sfrmPartDashboardDates.allowStepEdit

On Error Resume Next
Me.currentGate = Nz(Me.sfrmPartDashboardDates.Controls("gateTitle"))
Me.currentStep = Nz(Me.sfrmPartDashboard.Controls("stepType"))
Me.currentStepResp = Nz(findDept(Me.partNumber, Me.sfrmPartDashboard.Controls("responsible"), True, True))
On Error GoTo Err_Handler

Me.tabIssues.Caption = " Issues (" & issueCount(Me.partNumber) & ")"

'Who is the owner?
Select Case DLookup("templateType", "tblPartProjectTemplate", "recordId = " & Me.projectTemplateId)
    Case 1 'New Model
        TempVars.Add "projectOwner", "Project"
    Case 2 'Service
        TempVars.Add "projectOwner", "Service"
End Select

If restrict(Environ("username"), TempVars!projectOwner) Then 'if NOT owner
    Form_sfrmPartDashboard.newStep.Visible = False
    Form_sfrmPartDashboard.btnEditInfo.Visible = False
    Me.projectCapitalNumber.Locked = True
    Form_sfrmPartDashboardDates.plannedDate.Locked = True
    Form_sfrmPartDashboardDates.actualDate.Locked = True
    Form_sfrmPartDashboard.dueDate.Locked = True
End If

Me.tabAutomation.Visible = Me.projectTemplateId = 8

Form_sfrmPartDashboard.userSelect = Environ("username")
Form_sfrmPartDashboard.userSelect.RowSource = "SELECT person FROM tblPartTeam WHERE partNumber = '" & Me.partNumber & "'"
Select Case userData("Level")
    Case "Supervisor", "Manager"
        Form_sfrmPartDashboard.userSelect.Visible = True
    Case Else
        Form_sfrmPartDashboard.userSelect.Visible = False
End Select

Dim assyNum As String
assyNum = Nz(DLookup("assemblyNumber", "tblPartComponents", "componentNumber = '" & partNum & "'"), "")
Me.openAssy.Visible = False
If assyNum <> "" Then
    Me.openAssy.Visible = True
    Me.openAssy.Caption = " Open " & assyNum
End If

Dim rsPartInfo As Recordset, rsProgram As Recordset
Set rsPartInfo = db.OpenRecordset("SELECT * FROM tblPartInfo WHERE partNumber = '" & Me.partNumber & "'")
If rsPartInfo.RecordCount = 0 Then GoTo noProgram
If Nz(rsPartInfo!programId, 0) = 0 Then GoTo noProgram
Set rsProgram = db.OpenRecordset("SELECT * FROM tblPrograms WHERE ID = " & rsPartInfo!programId)
If rsProgram.RecordCount = 0 Then GoTo noProgram

'checks PASSED, program found
Me.modelCode = rsProgram!modelCode
Me.CarPicture.Picture = Nz(rsProgram!CarPicture, "")
    
rsProgram.CLOSE
Set rsProgram = Nothing

noProgram:

On Error Resume Next
rsPartInfo.CLOSE
Set rsPartInfo = Nothing
rsProjects.CLOSE
Set rsProjects = Nothing
Set db = Nothing

Me.tabSteps.SetFocus

Exit Sub
killDash:
DoCmd.CLOSE acForm, "frmPartDashboard"
Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub partInfo_Click()
On Error GoTo Err_Handler

Form_DASHBOARD.partNumberSearch = Me.partNumber.Value

If DCount("recordId", "tblPartProjectPartNumbers", "projectId = " & Me.recordId) > 0 Then
    TempVars.Add "partDashAction", "frmPartInformation"
    DoCmd.OpenForm "frmPartProjectPartNumbers", acNormal, , "[projectId] = " & Me.recordId
    Form_frmPartProjectPartNumbers.doAction.Visible = True
    Form_frmPartProjectPartNumbers.doActionMaster.Visible = True
Else
    TempVars.Add "partNumber", Me.partNumber.Value
    DoCmd.OpenForm "frmPartInformation"
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub projectCapitalNumber_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartProject", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.recordId)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Public Sub partDash_refresh_Click()
On Error GoTo Err_Handler

Dim gateId As Long
gateId = Form_sfrmPartDashboardDates.recordId

Form_frmPartDashboard.tabIssues.Caption = " Issues (" & issueCount(Form_frmPartDashboard.partNumber) & ")"
Form_frmPartDashboard.Requery

With Form_sfrmPartDashboardDates.RecordsetClone
    .FindFirst "recordId = " & gateId
    If Not .noMatch Then
        Form_sfrmPartDashboardDates.Bookmark = .Bookmark
    End If
End With

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub reports_Click()
On Error GoTo Err_Handler

Form_DASHBOARD.partNumberSearch = Me.partNumber.Value
DoCmd.OpenForm "frmReports"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
