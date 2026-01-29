Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub addProgram_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmAddProgram"

Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub Adjusted_Due_Date_AfterUpdate()
On Error GoTo Err_Handler
If IsNull(Me.Control_Number) = False Then
    Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
End If
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub cboAdjustedReason_AfterUpdate()
On Error GoTo Err_Handler
If IsNull(Me.Control_Number) = False Then Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub cboApprovalStatus_AfterUpdate()
On Error GoTo Err_Handler
If IsNull(Me.Control_Number) = False Then Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)

If Me.cboApprovalStatus = 2 Then        ' approved
    Me.Requester = Environ("username")
    Call registerDRSUpdates("tblDRS", Me.Control_Number, "Requester", Me.Requester.OldValue, Me.Requester)
End If
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Function checkUserDuplicates()
checkUserDuplicates = True

Select Case True
    Case Nz(Me.Assignee, "as") = Nz(Me.Checker_1, "ck")
        MsgBox "Can't have assignee and checker be the same!", vbExclamation, "No can do"
        checkUserDuplicates = False
    Case Nz(Me.Assignee, "as") = Nz(Me.Checker_2, "ap")
        MsgBox "Can't have assignee and approver be the same!", vbExclamation, "No can do"
        checkUserDuplicates = False
    Case Nz(Me.Checker_1, "ck") = Nz(Me.Checker_2, "ap")
        MsgBox "Can't have checker and approver be the same!", vbExclamation, "No can do"
        checkUserDuplicates = False
End Select

End Function

Private Sub cboAssignee_AfterUpdate()
On Error GoTo Err_Handler
If checkUserDuplicates = False Then
    Me.ActiveControl = Me.ActiveControl.OldValue
    Exit Sub
End If
If IsNull(Me.Control_Number) = False Then
    Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
End If
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub cboChecker1_AfterUpdate()
On Error GoTo Err_Handler
If checkUserDuplicates = False Then
    Me.ActiveControl = Me.ActiveControl.OldValue
    Exit Sub
End If
If IsNull(Me.Control_Number) = False Then
    Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
End If
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub cboChecker2_AfterUpdate()
On Error GoTo Err_Handler
If checkUserDuplicates = False Then
    Me.ActiveControl = Me.ActiveControl.OldValue
    Exit Sub
End If
If IsNull(Me.Control_Number) = False Then
    Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
End If
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub cboComplexity_AfterUpdate()
On Error GoTo Err_Handler
If IsNull(Me.Control_Number) = False Then
    Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
End If
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub cboDelayReason_AfterUpdate()
On Error GoTo Err_Handler
If IsNull(Me.Control_Number) = False Then
    Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
End If
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub cboDesignLevel_AfterUpdate()
On Error GoTo Err_Handler
If IsNull(Me.Control_Number) = False Then
    Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
End If
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub cboDesignResponsibility_AfterUpdate()
On Error GoTo Err_Handler

If IsNull(Me.Control_Number) = False Then
    Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
End If

Dim bDFMEA As Boolean
bDFMEA = DLookup("drs_dfmea", "tblDropDownsSP", "[recordid] = " & Nz(Me.cboDesignResponsibility, 3))
    If bDFMEA = True Then
        Me.lblDFMEA.Caption = "DFMEA"
        Me.lblDFMEA.Visible = True
    Else
        Me.lblDFMEA.Caption = ""
        Me.lblDFMEA.Visible = False
    End If
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub cboDRLevel_AfterUpdate()
On Error GoTo Err_Handler
If IsNull(Me.Control_Number) = False Then
    Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
End If
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub cboModelCode_AfterUpdate()
On Error GoTo Err_Handler
If IsNull(Me.Control_Number) = False Then
    Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
End If
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub cboModelCode_NotInList(newData As String, response As Integer)
On Error GoTo Err_Handler

MsgBox "Please click Add Program button to add a new program.", vbOKOnly, "Oops"

Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub cboProjectLocation_AfterUpdate()
On Error GoTo Err_Handler
If IsNull(Me.Control_Number) = False Then
    Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
End If
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub cboRequestType_AfterUpdate()
On Error GoTo Err_Handler

setVis_DRSLoc

If Me.cboRequestType = 10 Or Me.cboRequestType = 4 Then
    Me.cboToolingDept.Locked = False
Else
    Me.cboToolingDept = Null
    Me.cboToolingDept.Locked = True
End If

If IsNull(Me.Control_Number) = False Then
    Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
End If
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub cboToolingDept_AfterUpdate()
On Error GoTo Err_Handler
If IsNull(Me.Control_Number) = False Then
    Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
End If
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub cmdAdd_Click()
On Error GoTo Err_Handler

DoCmd.GoToRecord , , acNewRec
Me.cboRequestType.SetFocus
Me.Repaint

Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub cmdClone_Click()
On Error GoTo Err_Handler
Dim db As DAO.Database
    Dim strSQL As String
    Dim strOldComments As String
    Dim lngNewID As Long, lngOldID As Long
    bClone = True
    Set db = CurrentDb()
    strOldComments = Me.Comments
    
    lngOldID = Me.Control_Number

If Me.Dirty = True Then Me.Dirty = False
     
If IsNull(Me.Comments) Then
    MsgBox "[Comments] is a required field.  Please complete your entry.", vbOKOnly, "Required Field"
    Me.Comments.SetFocus
    Exit Sub
Else
'-- duplicate the main record: add to form's clone.
    With Me.RecordsetClone
        .addNew
            !Issue_Date = Date
            !Requester = Me.Requester
            !DR_Level = Me.DR_Level
            !Request_Type = Me.Request_Type
            !Part_Number = Me.Part_Number
            !PART_DESCRIPTION = Me.PART_DESCRIPTION
            !DESIGN_RESPONSIBILITY = Me.DESIGN_RESPONSIBILITY
            !Model_Code = Me.Model_Code
            !Part_Complexity = Me.Part_Complexity
            !Design_Level = Me.Design_Level
            !Assignee = Me.Assignee
            !Request_Number = Me.Request_Number
            !Due_Date = Me.Due_Date
            !Checker_1 = Me.Checker_1
            !Checker_2 = Me.Checker_2
            !Project_Location = Me.Project_Location
            !Tooling_Department = Me.Tooling_Department
            !Adjusted_Due_Date = Me.Adjusted_Due_Date
            !Adjusted_Reason = Me.Adjusted_Reason
            !Delay_Reason = Me.Delay_Reason
            !Approval_Status = Me.Approval_Status
        .Update
        
'-- save the primary key value, to use as the foreign key for the related record
        .Bookmark = .lastModified
        lngNewID = !Control_Number
        
'-- duplicate the related records: append query.
            strSQL = "INSERT INTO [dbo_tblComments] (Control_Number, Comments ) VALUES (" & lngNewID & ", '" & StrQuoteReplace(strOldComments) & "');"
            db.Execute strSQL, dbFailOnError
        Me.Requery
    End With
    
    Me.FilterOn = False
    Me.Recordset.FindFirst "Control_Number = " & lngNewID
End If
    Set db = Nothing
    MsgBox "You are now on the cloned copy of the record.  Please modify as needed.", vbOKOnly, "Record Cloned"
    Call registerDRSUpdates("tblDRS", Me.Control_Number, "DRS creation", "", "New DRS Cloned from DRS#" & lngOldID)
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub cmdFilter_Click()
On Error GoTo Err_Handler
Dim strFormName As String
Dim strSQL As String
    
    strFormName = "frmFilter"
    strSQL = "SELECT * FROM dbo_tblDRS WHERE [Approval_Status] <>3 ORDER BY [Control_Number] DESC;"
    
    If Me.cmdFilter.ControlTipText = "Remove Filter" Then
        Me.RecordSource = strSQL
        Me.Requery
        Me.lblFiltered.Visible = False
'-- toggle filter button properties
        Me.cmdFilter.ControlTipText = "Apply Filter"
    Else
        DoCmd.OpenForm strFormName, acNormal
    End If
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub cmdFind_Click()
On Error GoTo Err_Handler
Me.allowEdits = False
DoCmd.RunCommand acCmdFind
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub cmdFirst_Click()
On Error GoTo Err_Handler
DoCmd.GoToRecord , , acFirst
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub cmdGetDNumber_Click()
On Error GoTo Err_Handler

If MsgBox("Once you pull a D number, it is reserved forever. Are you sure?", vbYesNo, "Just making sure") <> vbYes Then Exit Sub

If Left(Me.Part_Number, 1) = "D" And Me.Part_Number = "D" & DMax("[dNumber]", "tblDnumbers") Then
    MsgBox "Are you trying to spam the system?", vbQuestion, "Looks like you already pulled one.."
    Exit Sub
End If

Me.Part_Number = createDnumber()

Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub cmdLast_Click()
On Error GoTo Err_Handler
DoCmd.GoToRecord , , acLast
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub cmdMail_Click()
On Error GoTo Err_Handler
Dim SendItems As New clsOutlookCreateItem               ' outlook class
    Dim strRequesttype As String                            ' type of request (for subject)
    Dim strFilePath As String                               ' path to archive
    Dim strFileName As String                               ' output file name
    Dim strReportName As String                             ' report name
    Dim strTo As String                                     ' email recipient
    Dim strSubject As String                                ' email subject
    Dim strBody As String                                   ' email body (text)
    Dim strAttach As String                                 ' attachment(s)
    
    If Me.Dirty = True Then Me.Dirty = False
    
    If Me.Approval_Status = 1 Then
        If MsgBox("This WO is still Pending, would you like to set it to Approved before you email?", vbYesNo, "WO Status is Pending") = vbYes Then
            Me.Approval_Status = 2
            If Me.Dirty = True Then Me.Dirty = False
            Me.refresh
            If IsNull(Me.Control_Number) = False Then Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.cboApprovalStatus.name, Me.cboApprovalStatus.OldValue, Me.cboApprovalStatus)
        End If
        If Me.cboApprovalStatus = 2 Then        ' approved
            Me.Requester = Environ("username")
            Call registerDRSUpdates("tblDRS", Me.Control_Number, "Requester", Me.Requester.OldValue, Me.Requester)
        End If
    End If
    
    Set SendItems = New clsOutlookCreateItem
    strRequesttype = DLookup("[drs_type]", "tblDropDownsSP", "[recordid] = " & Me.Request_Type)
    strFilePath = "\\cifs\PDFrepos\db\PreProdData\archive\"
    strFileName = Me.Control_Number & " DRS for " & Me.Part_Number & " " & strRequesttype & " Due " & Format(Me.Due_Date, "mmddyyyy") & ".pdf"
    strReportName = "rptDesignRequest"
    If IsNull(Me.Assignee) Then
        MsgBox "You can't send an email without an Assignee.", vbOKOnly, "No Assignee"
        Exit Sub
    Else
        strTo = getEmail(DLookup("[user]", "tblPermissions", "[ID] = " & Me.Assignee))
    End If
    strSubject = Me.Control_Number & " DRS for " & Me.Part_Number & " " & strRequesttype & " Due " & Me.Due_Date
    strBody = Me.Comments
    strAttach = strFilePath & strFileName
    
'-- make sure user has entered comments
    If IsNull(Me.Comments) Then
        MsgBox "[Comments] is a required field.  Please complete your entry.", vbOKOnly, "Required Field"
        Me.Comments.SetFocus
        Exit Sub
    Else
    
'-- generate report
    'DoCmd.OpenReport strReportName, acViewPreview, , "[Control_Number]=" & Me.Control_Number, acHidden
    'DoCmd.OutputTo acOutputReport, strReportName, acFormatPDF, strFilePath & strFileName, False
    DoCmd.CLOSE acReport, strReportName
    
'-- create and send mail object
    SendItems.CreateMailItem sendTo:=strTo, _
                             subject:=strSubject, _
                             body:=strBody, _
                             Attachments:="" 'strAttach
    End If
    
'-- housekeeping
    Set SendItems = Nothing
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub cmdMailApprover_Click()
On Error GoTo Err_Handler

Call save_Click

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub cmdNext_Click()
On Error GoTo Err_Handler

DoCmd.GoToRecord , , acNext
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub cmdPreview_Click()
On Error GoTo Err_Handler

Dim strReportName As String
Dim strLinkCriteria As String

If Me.Dirty = True Then Me.Dirty = False
    
    strReportName = "rptDesignRequest"
    strLinkCriteria = "[Control_Number] = " & Me.Control_Number
    
    If IsNull(Me.Comments) Then
        MsgBox "[Comments] is a required field.  Please complete your entry.", vbOKOnly, "Required Field"
        Me.Comments.SetFocus
        Exit Sub
    Else
    DoCmd.OpenReport strReportName, acViewPreview, , strLinkCriteria
    End If
    
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo Err_Handler

DoCmd.GoToRecord , , acPrevious
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Comments_AfterUpdate()
On Error GoTo Err_Handler
Call registerDRSUpdates("tblComments", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Completed_Date_AfterUpdate()
On Error GoTo Err_Handler
If IsNull(Me.Control_Number) = False Then
    Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub drsHistory_Click()
On Error GoTo Err_Handler
If IsNull(Me.Control_Number) = False Then
    DoCmd.OpenForm "frmHistory"
    Form_frmHistory.RecordSource = "qryDRSupdateTracking"
    Form_frmHistory.dataTag0.ControlSource = "dataTag0"
    Form_frmHistory.previousData.ControlSource = "previous"
    Form_frmHistory.newData.ControlSource = "new"
    Form_frmHistory.filter = "tableRecordId = " & Me.Control_Number
    Form_frmHistory.FilterOn = True
    Form_frmHistory.OrderBy = "updatedDate Desc"
    Form_frmHistory.OrderByOn = True
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Due_Date_AfterUpdate()
On Error GoTo Err_Handler
If IsNull(Me.Control_Number) = False Then
    Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub etaSpread_Click()
On Error GoTo Err_Handler

populateWorkload
Call populateETAs(Me.Issue_Date, Nz(Me.Adjusted_Due_Date, Me.Due_Date))
DoCmd.OpenForm "frmTimeViewETA"

Form_frmTimeViewETA.issueDate = Me.Issue_Date
Form_frmTimeViewETA.dueDate = Nz(Me.Adjusted_Due_Date, Me.Due_Date)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Current()
On Error GoTo Err_Handler
Dim rs As DAO.Recordset
Dim iFcst As Integer
Dim iLevel As Integer
Dim lngCount As Long
Dim bDFMEA As Boolean
Dim strSQL1 As String
Dim dJudgeDate As Date
Dim dCompDate As Date
    
    Set rs = Me.RecordsetClone
    With rs
        .MoveFirst
        .MoveLast
        lngCount = .RecordCount
    End With
   
    iLevel = Nz(DLookup("[designWOpermissions]", "tblPermissions", "[user] = '" & Environ("username") & "'"), 3)
   
'-- set record counter values
    Me.txtXofY = "Record " & Me.CurrentRecord & " of " & lngCount
    
'-- customize behavior of navigation buttons
    Call SetNavButtons
    
    If Me.NewRecord Then
        bDFMEA = False
        strSQL1 = "SELECT [ID], [lastName] & ', ' & [firstName] AS EngName FROM tblPermissions WHERE [Inactive] = False AND designWOpermissions <> 3 ORDER BY [lastName];"
    Else
        bDFMEA = DLookup("drs_dfmea", "tblDropDownsSP", "[recordid] = " & Nz(Me.cboDesignResponsibility, 3))
        dJudgeDate = IIf(IsNull(Me.Adjusted_Due_Date), Me.Due_Date, Me.Adjusted_Due_Date)
        dCompDate = IIf(IsNull([Completed_Date]), Date, [Completed_Date])
        If Me.Approval_Status = 1 Then
            strSQL1 = "SELECT [ID], [lastName] & ', ' & [firstName] AS EngName FROM tblPermissions WHERE [Inactive] = False AND designWOpermissions <> 3 ORDER BY [lastName];"
        Else
            strSQL1 = "SELECT [ID], [lastName] & ', ' & [firstName] AS EngName FROM tblPermissions WHERE designWOpermissions <> 3 ORDER BY [lastName];"
        End If
    End If
    
'-- if design responsibility is indicated for dfmea, flag it as such
    If bDFMEA = True Then
        Me.lblDFMEA.Caption = "DFMEA"
        Me.lblDFMEA.Visible = True
    Else
        Me.lblDFMEA.Caption = ""
        Me.lblDFMEA.Visible = False
    End If
    
    setVis_DRSLoc
    
'-- assign row source to various staff-related fields
    Me.cboAssignee.RowSource = strSQL1
    Me.cboChecker1.RowSource = strSQL1
    Me.cboChecker2.RowSource = strSQL1
    
'-- judgment label
    If dCompDate <= dJudgeDate Then
        Me.lblJudgment.Caption = "On Time"
    Else
        Me.lblJudgment.Caption = "Late"
    End If
    
'-- just in case user navigated here using the find button, we want to allow edits (which was turned off to hide the replace tab in the find dialog)
    Me.allowEdits = True
    Me.Repaint
'-- housekeeping
    rs.CLOSE
    Set rs = Nothing
    
    Exit Sub
Err_Handler:
    If Err.number = 3021 Then
        MsgBox "Nothing found", vbInformation, "No Records"
        Me.FilterOn = False
        Exit Sub
    End If
    Call handleError(Me.name, "Form_Current", Err.DESCRIPTION, Err.number)
End Sub

Function setVis_DRSLoc()
On Error GoTo Err_Handler

Dim showIt As Boolean

showIt = Nz(Me.Request_Type, 0) = 19 'only show for customer meetings
Me.cboDRSLocation.Visible = showIt
Me.Label132.Visible = showIt
Me.Command130.Visible = showIt

Exit Function
Err_Handler:
    Call handleError(Me.name, "setVis_DRSLoc", Err.DESCRIPTION, Err.number)
End Function

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.cboRequestType.SetFocus

Dim iLevel As Integer
Dim strSQL1 As String
Dim ctl As Control

    iLevel = Nz(DLookup("[designWOpermissions]", "tblPermissions", "[user] = '" & Environ("username") & "'"), 3)
    strSQL1 = "SELECT [ID], [lastName] & ', ' & [firstName] AS EngName FROM tblPermissions WHERE [Inactive] = False AND designWOpermissions <> 3 ORDER BY [lastName];"
    
    bClone = False
    
    Me.cboAssignee.DefaultValue = DLookup("[ID]", "[tblPermissions]", "[user] = '" & Environ("username") & "'")
    Me.Requester.DefaultValue = "'" & Environ("username") & "'"
    
    Me.cboAssignee.RowSource = strSQL1
    Me.cboChecker1.RowSource = strSQL1
    Me.cboChecker2.RowSource = strSQL1

Select Case iLevel
        Case Is = 1         ' manager
            For Each ctl In Me.Controls
                If ctl.tag Like "*mgrVis*" Then ctl.Visible = True
            Next
        Case Is = 2         ' initiator
            For Each ctl In Me.Controls
                If ctl.tag Like "*mgrVis*" Then ctl.Visible = False
            Next
            
        Case Is = 3         ' read only
            Me.AllowAdditions = False
            Me.AllowDeletions = False
            Me.allowEdits = False
            Me.cmdClone.Enabled = False
            Me.cmdAdd.Enabled = False
            Me.save.Enabled = False

            For Each ctl In Me.Controls
                If ctl.tag Like "*mgrVis*" Then ctl.Visible = False
            Next
            For Each ctl In Me.Controls
                If ctl.tag Like "*?*" Then ctl.Locked = True
            Next
        Case Else
    End Select
    
If TempVars!drsNewRec = True Then
    DoCmd.GoToRecord , , acNewRec
    TempVars.Add "drsNewRec", "False"
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Sub SetNavButtons()
On Error Resume Next
'-- enable/disable buttons depending on record position

If Nz(DLookup("[designWOpermissions]", "tblPermissions", "[user] = '" & Environ("username") & "'"), 3) <> 1 Then Exit Sub 'only managers

Me.cmdFirst.Enabled = False
Me.cmdPrevious.Enabled = False
Me.cmdNext.Enabled = False
Me.cmdLast.Enabled = False
Select Case True
    Case Me.Recordset.RecordCount <= 1 Or Me.CurrentRecord > Me.Recordset.RecordCount
        Me.cmdFirst.Enabled = True
        Me.cmdPrevious.Enabled = True
        Me.cmdPrevious.SetFocus
    Case Me.CurrentRecord = 1
        Me.cmdNext.Enabled = True
        Me.cmdLast.Enabled = True
        Me.cmdNext.SetFocus
    Case Else
        Me.cmdFirst.Enabled = True
        Me.cmdPrevious.Enabled = True
        Me.cmdNext.Enabled = True
        Me.cmdLast.Enabled = True
End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler

If validate <> "" Then
    If MsgBox("Are you sure?" & vbNewLine & "Your current record will be deleted.", vbYesNo, "Please confirm") <> vbYes Then
        Cancel = True
        Exit Sub
    End If

    DoCmd.SetWarnings False
    If Nz(Me.Control_Number) <> "" Then DoCmd.RunCommand acCmdDeleteRecord
    DoCmd.SetWarnings True
End If

If CurrentProject.AllForms("frmDRSworkTracker").IsLoaded = True Then Form_frmDRSworkTracker.refresh_Click

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub imgAssignee_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmUserProfile", , , "user = '" & Me.userAssignee & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub imgChecker1_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmUserProfile", , , "user = '" & Me.userChecker1 & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub imgChecker2_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmUserProfile", , , "user = '" & Me.userChecker2 & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub imgRequester_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmUserProfile", , , "user = '" & Me.Requester & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Part_Description_AfterUpdate()
On Error GoTo Err_Handler
If IsNull(Me.Control_Number) = False Then
    Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Part_Number_AfterUpdate()
On Error GoTo Err_Handler

If IsNull(Me.Control_Number) = False Then
    Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
End If

Dim partNum As String
partNum = Me.Part_Number

Dim errorTxt As String
errorTxt = ""

If Len(partNum) > 5 Then
    'check for P, D, Q part numbers then don't allow anything else
    
    'first, look for NCM part numbers. kill those.
    If partNum Like "[A-Z][A-Z]##[A-Z]##[A-Z]" Or partNum Like "[A-Z][A-Z]##[A-Z]##" Then
        errorTxt = "Please only enter the 5 character part number. For NCM part numbers, just use '12A36' instead of 'AB12A36A'"
        GoTo errorCheck
    End If
    
    'then look for P, D, Q numbers. allow those.
    If Not (partNum Like "*P" Or partNum Like "D*" Or partNum Like "Q*") Then
        errorTxt = "Please only enter 5 characters for a part number."
        GoTo errorCheck
    End If
End If

errorCheck:
If errorTxt <> "" Then
    MsgBox errorTxt, vbInformation, "Notice!"
    Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.ActiveControl.name, Me.Part_Number, "")
    Me.Part_Number = ""
    GoTo exitThis
End If

If DCount("[Part_Number]", "dbo_tblDRS", "[Part_Number] = '" & partNum & "'") > 0 Then
    Me.PART_DESCRIPTION = DLookup("[Part_Description]", "dbo_tblDRS", "[Part_Number] = '" & partNum & "'")
    Me.DESIGN_RESPONSIBILITY = DLookup("[Design_Responsibility]", "dbo_tblDRS", "[Part_Number] = '" & partNum & "'")
    Me.Model_Code = DLookup("[Model_Code]", "dbo_tblDRS", "[Part_Number] = '" & partNum & "'")
    Me.Part_Complexity = DLookup("[Part_Complexity]", "dbo_tblDRS", "[Part_Number] = '" & partNum & "'")
    Me.Project_Location = DLookup("[Project_Location]", "dbo_tblDRS", "[Part_Number] = '" & partNum & "'")
    Me.Tooling_Department = DLookup("[Tooling_Department]", "dbo_tblDRS", "[Part_Number] = '" & partNum & "'")
    GoTo exitThis
End If

If Left(partNum, 1) = "D" Then
    If MsgBox("I can't find this D number.  Do you want to create a new 'D' Number?", vbYesNo, "No Matching Part") = vbYes Then Call cmdGetDNumber_Click
    GoTo exitThis
End If

On Error Resume Next
If DCount("[Description]", "APPS_MTL_SYSTEM_ITEMS", "[SEGMENT1] = '" & partNum & "'") > 0 Then
    Me.PART_DESCRIPTION = Nz(DLookup("[Description]", "APPS_MTL_SYSTEM_ITEMS", "[SEGMENT1] = '" & partNum & "'"))
    GoTo exitThis
End If

If DCount("newPartNumber", "tblPartNumbers", "newPartNumber = " & partNum) > 0 Then 'NEW PART LOG
    Me.PART_DESCRIPTION = Nz(DLookup("[partDescription]", "tblPartNumbers", "[newPartNumber] = '" & partNum & "'"))
    GoTo exitThis
End If

If DCount("[Nifco_Part_Number]", "qryUnionPartDescriptions", "[Nifco_Part_Number] = '" & partNum & "'") > 0 Then   'SIFS
    Me.PART_DESCRIPTION = Nz(DLookup("[Part_Description]", "qryUnionPartDescriptions", "[Nifco_Part_Number] = '" & partNum & "'"))
    GoTo exitThis
End If

exitThis:
Me.refresh
    
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Request_Number_AfterUpdate()
On Error GoTo Err_Handler
If IsNull(Me.Control_Number) = False Then Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function validate() As String
validate = ""

Select Case True
    Case Nz(Me.Request_Type) = ""
        validate = "Request Type"
    Case Nz(Me.Design_Level) = ""
        validate = "ETA"
    Case Nz(Me.Due_Date) = ""
        validate = "Due Date"
    Case Nz(Me.Part_Number) = ""
        validate = "Part Number"
    Case Nz(Me.PART_DESCRIPTION) = ""
        validate = "Part Description"
    Case Nz(Me.Comments) = ""
        validate = "Comments"
End Select

If Me.Request_Type = 19 Then
    If Nz(Me.DRS_Location, 0) = 0 Then
        validate = "Meeting Type"
    End If
End If

End Function

Private Sub resetChecksheet_Click()
On Error GoTo Err_Handler

If MsgBox("Are you sure? This will delete all checksheet items for this WO, including all comments.", vbYesNo, "Are you sure?") = vbYes Then
    Dim db As Database
    Set db = CurrentDb()
    Dim rsChecksheet As Recordset
    Set rsChecksheet = db.OpenRecordset("SELECT * FROM tblDesignChecksheet WHERE controlNumber = " & Me.Control_Number)
    
    Do While Not rsChecksheet.EOF
        rsChecksheet.Delete
        rsChecksheet.MoveNext
    Loop
    
    rsChecksheet.CLOSE
    Set rsChecksheet = Nothing
    Set db = Nothing
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub save_Click()
On Error GoTo Err_Handler

If Me.Dirty = True Then Me.Dirty = False

Dim val As String
val = validate
If val <> "" Then
    MsgBox "Must Enter " & val, vbInformation, "Fix it"
    Exit Sub
End If

If (DCount("Control_Number", "tblDRStrackerExtras", "Control_Number = " & Me.Control_Number) = 0) Then
    dbExecute "INSERT INTO tblDRStrackerExtras(Control_Number,Check_In_Prog) VALUES (" & Me.Control_Number & ",'Not Started')"
    Call registerDRSUpdates("tblDRS", Me.Control_Number, "DRS creation", "", "New DRS")
End If

If Nz(DLookup("[designWOpermissions]", "tblPermissions", "[user] = '" & Environ("username") & "'"), 3) = 1 Then 'manager
    If Me.Approval_Status <> 1 Then GoTo notPending
    If MsgBox("This WO is still Pending, would you like to set it to Approved before saving?", vbYesNo, "WO Status is Pending") = vbYes Then
        Me.Approval_Status = 2
        If Me.Dirty Then Me.Dirty = False
        Me.refresh
        
        If Not IsNull(Me.Control_Number) Then Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.cboApprovalStatus.name, Me.cboApprovalStatus.OldValue, Me.cboApprovalStatus)
        
        Me.Requester = Environ("username")
        Call registerDRSUpdates("tblDRS", Me.Control_Number, "Requester", Me.Requester.OldValue, Me.Requester)
    End If
Else 'if NOT manager, try to send email
    If Nz(Me.cmbUsername) = "" Then
        MsgBox "You gotta enter a username of someone to email!"
        Exit Sub
    End If
    GoTo sendApproverEmail
End If

If Nz(Me.cmbUsername) <> "" Then GoTo sendApproverEmail

GoTo notPending
sendApproverEmail:
Dim strSubject As String, strBody As String
strSubject = "DRS Request: " & Me.Control_Number & " for PN " & Me.Part_Number
strBody = "<body>Hi " & userData("firstName", Me.cmbUsername) & "," & "<br/>" & "<br/>" & "Please approve this DRS: " & "<br/>" & _
            "DRS#" & Me.Control_Number & " " & Me.cboRequestType.column(1) & " for " & Me.Part_Number & ": " & Me.Comments _
            & "<br/>" & "<br/>" & "Thank you," & "<br/>" & userData("firstName") & "</body>"
Call wdbEmail(getEmail(Me.cmbUsername), "", strSubject, strBody)

notPending:

If IsNull(Me.Control_Number) = False Then
    If IsNull(Me.Comments) Then
        MsgBox "[Comments] is a required field.  Please complete your entry.", vbOKOnly, "Required Field"
        Me.Comments.SetFocus
        Exit Sub
    End If
    If DCount("Control_Number", "tblDRStrackerExtras", "Control_Number = " & Me.Control_Number) = 0 Then dbExecute "INSERT INTO tblDRStrackerExtras(Control_Number,Check_In_Prog) VALUES (" & Me.Control_Number & ",'Not Started')"
End If

If CurrentProject.AllForms("frmDRSworkTracker").IsLoaded = True Then Form_frmDRSworkTracker.refresh_Click
Me.Form.SetFocus
DoCmd.CLOSE acForm, "frmApproveDRS"
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
