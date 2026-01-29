Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub attachHistory_Click()
On Error GoTo Err_Handler

Select Case True
    Case IsNull(Me.TtestId) And IsNull(Me.TstepId) 'all attachment history
        DoCmd.OpenForm "frmHistory", acNormal, , "[tableName] = 'tblPartAttachmentsSP' AND [partNumber] = '" & Me.TpartNumber & "'"
    Case Nz(Me.secondaryType) = 22 'test attachment history
        DoCmd.OpenForm "frmHistory", acNormal, , "[tableName] = 'tblPartAttachmentsSP' AND [partNumber] = '" & Me.TpartNumber & "' AND [tableRecordId] = " & Me.TtestId & " AND columnName = 'Testing Attachment'"
    Case Nz(Me.secondaryType) = 32 'trial attachment history
        DoCmd.OpenForm "frmHistory", acNormal, , "[tableName] = 'tblPartAttachmentsSP' AND [partNumber] = '" & Me.TpartNumber & "' AND [tableRecordId] = " & Me.TstepId & " AND columnName = 'Trial Attachment'"
    Case Else
        DoCmd.OpenForm "frmHistory", acNormal, , "[tableName] = 'tblPartAttachmentsSP' AND [partNumber] = '" & Me.TpartNumber & "' AND [tableRecordId] = " & Me.TstepId & " AND columnName = 'Step Attachment'"
End Select

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Detail_Paint()
On Error Resume Next

Me.openAttachment.Transparent = Me.fileStatus <> "Uploaded"

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler
If CurrentProject.AllForms("frmPartDashboard").IsLoaded Then Form_sfrmPartDashboard.fileCountCaption.Requery
If CurrentProject.AllForms("frmDropFile").IsLoaded Then DoCmd.CLOSE acForm, "frmDropFile"

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Unload", Err.DESCRIPTION, Err.number)
End Sub

Private Sub imgUploadedBy_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmUserProfile", , , "user = '" & Me.uploadedBy & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub newAttachment_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmDropFile"
Form_frmDropFile.customName.Visible = False
Form_frmDropFile.Label62.Visible = False
Form_frmDropFile.Command63.Visible = False

Form_frmDropFile.TpartNumber = Me.TpartNumber
Form_frmDropFile.TprojectId = Me.TprojectId

Form_frmDropFile.TtestId = Nz(Me.TtestId, "")
Form_frmDropFile.TstepId = Nz(Me.TstepId, "")

If IsNull(Me.TtestId) Then
    Dim docType
    docType = DLookup("documentType", "tblPartSteps", "recordId = " & Nz(Me.TstepId, 0))
    If Nz(docType, 0) <> 0 Then
        Form_frmDropFile.documentType = docType
        Form_frmDropFile.documentType.Locked = True
    Else
        Form_frmDropFile.documentType = 30
        Form_frmDropFile.documentType.Locked = True
        Form_frmDropFile.customName.Visible = True
        Form_frmDropFile.Label62.Visible = True
        Form_frmDropFile.Command63.Visible = True
        Form_frmDropFile.customName = DLookup("stepType", "tblPartSteps", "recordId = " & Nz(Me.TstepId, 0))
    End If
Else 'If it's not attached to a step
    Form_frmDropFile.documentType = Me.secondaryType
    Form_frmDropFile.documentType.Locked = True
    Form_frmDropFile.TtestType = Me.itemName.Caption
End If

'if related parts found on this project, check what part number to open
Form_DASHBOARD.partNumberSearch = Me.TpartNumber.Value

If DCount("recordId", "tblPartProjectPartNumbers", "projectId = " & Nz(Me.TprojectId, 0)) > 0 Then
    TempVars.Add "partDashAction", "frmDropFile"
    DoCmd.OpenForm "frmPartProjectPartNumbers", acNormal, , "[projectId] = " & Nz(Me.TprojectId, 0)
    Form_frmPartProjectPartNumbers.doAction.Visible = True
    Form_frmPartProjectPartNumbers.doActionMaster.Visible = True
Else
    TempVars.Add "partNumber", Me.TpartNumber.Value
    DoCmd.OpenForm "frmDropFile"
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openAttachment_Click()
On Error GoTo Err_Handler

If Me.fileStatus = "Uploaded" Then
    Application.FollowHyperlink Me.directLink
    
    If IsNull(Me.TtestId) Or Me.TtestId = "" Then
        Call registerPartUpdates("tblPartAttachmentsSP", Me.ID, "Step Attachment", Me.attachName, "Opened", Me.partNumber, "Opened File", Form_frmPartDashboard.recordId)
    Else
        Call registerPartUpdates("tblPartAttachmentsSP", Me.ID, "Test Attachment", Me.attachName, "Opened", Me.partNumber, "Opened File", Form_frmPartDashboard.recordId)
    End If
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub remove_Click()
On Error GoTo Err_Handler

If Me.fileStatus <> "Uploaded" Then
    MsgBox "File must be fully uploaded in order to delete.", vbInformation, "Wait a second..."
    Exit Sub
End If

'TEMPORARY RESTRICTION OVERRIDE
'Project engineers can delete any file while in beta testing
If IsNull(TempVars!projectOwner) = False Then
    If restrict(Environ("username"), TempVars!projectOwner) = False Then GoTo theCorrectFellow 'is the bro a PE?
End If

'FIRST: are you the right person for the job???
Select Case userData("Dept")
    Case "Design" 'depts in Designdev
        If Me.businessArea = "Designdev" Then GoTo theCorrectFellow
    Case "Project", "Automation", "Tooling" 'depts in newmodelengineering
        If Me.businessArea = "newmodelengineering" Then GoTo theCorrectFellow
    Case "New Model Quality" 'depts in quality2
        If Me.businessArea = "quality2" Then GoTo theCorrectFellow
    Case "Processing" 'depts in manufacturing
        If Me.businessArea = "manufacturing" Then GoTo theCorrectFellow
End Select

MsgBox "You aren't the right person for the job. This document is in the " & Me.businessArea & " business area", vbCritical, "Darnit"
Exit Sub

theCorrectFellow:

If MsgBox("Are you sure?", vbYesNo, "Please confirm") <> vbYes Then Exit Sub

If IsNull(Me.TtestId) Or Me.TtestId = "" Then
    Call registerPartUpdates("tblPartAttachmentsSP", Me.TstepId, "Step Attachment", Me.attachName, "Deleted", Me.TpartNumber, Me.itemName.Caption, Me.partProjectId)
Else
    Call registerPartUpdates("tblPartAttachmentsSP", Me.TtestId, "Test Attachment", Me.attachName, "Deleted", Me.TpartNumber, Me.itemName.Caption, Me.partProjectId)
End If

Me.fileStatus = "Deleting"
If Me.Dirty Then Me.Dirty = False

Me.Requery
Me.refresh

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
