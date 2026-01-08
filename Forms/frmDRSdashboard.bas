Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub approvalDate_Click()
On Error GoTo Err_Handler

Me.Completed_Date = Me.checker2SignDate

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnCancel_Click()
On Error GoTo Err_Handler

Select Case Environ("username")
    Case Me.Assignee, Me.Checker1, Me.Checker2
    Case Else
        MsgBox "You need to be a part of this WO to cancel it.", vbCritical, "Nope."
        Exit Sub
End Select

If MsgBox("Are you sure you want to cancel this DRS? Gotta make sure.", vbYesNo, "Warning") <> vbYes Then Exit Sub

Me.Delay_Reason = 11
Call registerDRSUpdates("tblDRS", Me.Control_Number, "Delay_Reason", Me.Delay_Reason.OldValue, Me.Delay_Reason)
Me.Completed_Date = Me.Due_Date
Me.Check_In_Prog = "Not Started"
Call registerDRSUpdates("tblDRS", Me.Control_Number, "Completed_Date", Me.Completed_Date.OldValue, Me.Completed_Date)
Me.Dirty = False
Dim db As Database
Set db = CurrentDb()
If (Me.Approval_Status = "Pending") Then db.Execute "UPDATE dbo_tblDRS SET Approval_Status = 2 WHERE Control_Number = " & Me.Control_Number
Set db = Nothing
DoCmd.CLOSE
Form_frmDRSworkTracker.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function getCheckFolder() As String
On Error GoTo Err_Handler

Dim chkFold As String, X, controlNum
chkFold = Nz(Me.Check_Folder)
controlNum = Me.Control_Number

getCheckFolder = ""

If Len(chkFold) < 3 Then
    'look for check folder
    chkFold = findCheckFolder
    
    If chkFold = "" Then
        If MsgBox("No check folder found yet. Would you like to add one?", vbYesNo, "No Folder Found") <> vbYes Then Exit Function
        chkFold = InputBox("Paste Link to Check Folder Here", "Add Check Folder Link")
        
        If chkFold = "" Then Exit Function
    End If
    
    chkFold = Replace(chkFold, "'", "''")
    chkFold = replaceDriveLetters(chkFold)
    
    If InStr(chkFold, "C:\Users\" & Environ("username") & "\Nifco America Corporation\") Then
        chkFold = Replace(chkFold, Environ("username"), "CUST_USER")
    End If
    
    Call registerDRSUpdates("tblDRStrackerExtras", controlNum, "Check_Folder", "", chkFold)
    Me.Check_Folder = chkFold
    If Me.Dirty Then Me.Dirty = False
End If

If InStr(Left(chkFold, 10), "file") Then
    chkFold = Replace(chkFold, "%20", " ")
    chkFold = Right(Left(chkFold, Len(chkFold) - 1), Len(chkFold) - 10)
End If

If InStr(chkFold, "C:\Users\CUST_USER\Nifco America Corporation\") Then
    chkFold = Replace(chkFold, "CUST_USER", Environ("username"))
End If

If chkFold <> "" Then getCheckFolder = addLastSlash(chkFold)

Exit Function
Err_Handler:
    Call handleError("frmDRSdashboard", "getCheckFolder", Err.DESCRIPTION, Err.number)
End Function

Private Sub btnCheckFolder_Click()
On Error GoTo Err_Handler
Dim chkFold As String
chkFold = getCheckFolder
If FolderExists(chkFold) Then
    FollowHyperlink chkFold
Else
    Call snackBox("error", "Hmm...", "Something is wrong with the check folder", Me.name)
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnECODetails_Click()
On Error GoTo Err_Handler

If Nz(Me.ECO) = "" Then
    Call snackBox("error", "No ECO", "Please enter an ECO number into the text box first!", "frmDRSdashboard")
    Exit Sub
End If

DoCmd.OpenForm "frmECOs", , , "[CHANGE_NOTICE] = '" & UCase(Me.ECO) & "'"
Form_frmECOs.ECOsrch = Me.ECO

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function findCheckFolder() As String
On Error GoTo Err_Handler

findCheckFolder = ""
Dim docHis As String, checkFolderTrial As String

docHis = openDocumentHistoryFolder(Me.Part_Number, False)
checkFolderTrial = addLastSlash(docHis) & "Misc\Drawing Packets\" & Me.Part_Number & "_00_DWG_PKT_1000_CHECK"
If FolderExists(checkFolderTrial) Then findCheckFolder = checkFolderTrial

Exit Function
Err_Handler:
    Call handleError(Me.name, "findCheckFolder", Err.DESCRIPTION, Err.number)
End Function

Private Sub btnEditChkFold_Click()
On Error GoTo Err_Handler
Dim chkFold As String, X As String

chkFold = Nz(Me.Check_Folder)
If chkFold = "" Or chkFold = "\" Then chkFold = findCheckFolder

X = InputBox("Paste Link to Check Folder Here", "Add Check Folder Link", chkFold)

If StrPtr(X) = 0 Then Exit Sub
If X = "" Then
    If MsgBox("Nothing entered. Would you like to clear your check folder?", vbYesNo, "You didn't type anything") = vbNo Then Exit Sub
End If

X = replaceDriveLetters(addLastSlash(X))

If InStr(X, "C:\Users\" & Environ("username") & "\Nifco America Corporation\") Then
    X = Replace(X, Environ("username"), "CUST_USER")
End If

Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, "Check_Folder", chkFold, X)

Me.Check_Folder = X
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub assigneeSign_Click()
On Error GoTo Err_Handler

If MsgBox("This will remove your signature and all signatures ahead of yours. Are you sure?", vbYesNo, "Be careful here...") = vbNo Then
    Me.ActiveControl = Me.ActiveControl.OldValue
    Exit Sub
End If

If Me.checker2Sign = True Then 'remove checker2 signature
    Me.checker2Sign = False
    Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, "checker2Sign", True, False)
    Me.checker2Sign.Locked = True
    Me.checker2SignDate = Null
    Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, "checker2SignDate", Me.checker2SignDate.OldValue, "")
    Me.checker2SignDate.Locked = True
    Me.Check_In_Prog.SetFocus
End If

If Me.checker1Sign = True Then 'remove checker1 signature
    Me.checker1Sign = False
    Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, "checker1Sign", True, False)
    Me.checker1Sign.Locked = True
    Me.checker1SignDate = Null
    Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, "checker1SignDate", Me.checker1SignDate.OldValue, "")
    Me.checker1SignDate.Locked = True
    Me.Check_In_Prog.SetFocus
End If

Me.assigneeSign.Locked = True 'remove assignee signature
Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, "assigneeSign", True, False)
Me.assigneeSignDate = Null
Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, "assigneeSignDate", Me.assigneeSignDate.OldValue, "")
Me.assigneeSignDate.Locked = True

Me.Check_In_Prog = "In Progress"
Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, "Check_In_Prog", Me.Check_In_Prog.OldValue, "In Progress")
Me.Check_In_Prog.Locked = False
Me.DRSactionButton.Enabled = True
Me.Check_In_Prog.SetFocus
Me.assigneeSign.Visible = False
Me.Repaint
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnImportPhoto_Click()
On Error GoTo Err_Handler

Form_DASHBOARD.partNumberSearch = Me.Part_Number
DoCmd.OpenForm "frmPartPicture"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub cfmm_Click()
On Error GoTo Err_Handler

'---Define Variables---
Dim db As Database
Set db = CurrentDb()

Dim rsPartInfo As Recordset, rs1 As Recordset, rsRelatedParts As Recordset
Dim partType, meetingId, partProjectId, partNumMaster
Dim quoteNumber As Long, sellingPrice As Double, quotedCost As Double, quoteId

'---Check for Meeting---
meetingId = Nz(DLookup("recordId", "tblPartMeetings", "meetingType = 1 AND partNum = '" & Me.Part_Number & "'"), "")
If meetingId = "" Then
    '---Find the part project---

    partProjectId = Nz(DLookup("recordId", "tblPartProject", "partNumber = '" & Me.Part_Number & "'"), "Null")
    If partProjectId = "Null" Then
        partNumMaster = DLookup("relatedPN", "sqryRelatedParts_ParentPNs", "primaryPN = '" & Me.Part_Number & "'")
        partProjectId = Nz(DLookup("recordId", "tblPartProject", "partNumber = '" & partNumMaster & "'"), "Null")
    End If
    db.Execute "INSERT INTO tblPartMeetings(dateOfMeeting,partNum,meetingType,partProjectId) VALUES ('" & Now() & "','" & Me.Part_Number & "',1," & partProjectId & ")"
    TempVars.Add "meetingId", db.OpenRecordset("SELECT @@identity")(0).Value
    meetingId = TempVars!meetingId
End If

'---CHECK FOR ALL DATA SETUP---
'-----------------------------------------------------

'---Check tblPartInfo---
TempVars.Add "newPartInfoId", Nz(DLookup("recordId", "tblPartInfo", "partNumber = '" & Me.Part_Number & "'"), 0)
If TempVars!newPartInfoId = 0 Then
    db.Execute "INSERT INTO tblPartInfo(partNumber,description,designResponsibility) VALUES('" & Me.Part_Number & "','" & Me.PART_DESCRIPTION & "','" & Me.DESIGN_RESPONSIBILITY & "')"
    TempVars.Add "newPartInfoId", db.OpenRecordset("SELECT @@identity")(0).Value
End If
Set rsPartInfo = db.OpenRecordset("SELECT * FROM tblPartInfo WHERE recordId = " & TempVars!newPartInfoId)

'---Check for Quote Info---
If Nz(rsPartInfo!quoteInfoId, 0) = 0 Then
    quoteNumber = 0
    sellingPrice = 0
    quotedCost = 0
    
    If DCount("[ROW_ID]", "APPS_Q_SIF_NEW_MOLDED_PART_V", "[NIFCO_PART_NUMBER] = '" & Me.Part_Number & "'") > 0 Then 'IF MOLDED
        Set rs1 = db.OpenRecordset("APPS_Q_SIF_NEW_MOLDED_PART_V", dbOpenSnapshot)
        rs1.FindLast "[NIFCO_PART_NUMBER] = '" & Me.Part_Number & "'"
        quoteNumber = Left(rs1!ENG_QUOTE_LOG_NUM, 4)
        sellingPrice = rs1!PIECE_PRICE
        quotedCost = rs1!INTERNAL_PART_COST
    ElseIf DCount("[ROW_ID]", "APPS_Q_SIF_NEW_ASSEMBLED_PART_V", "[NIFCO_PART_NUMBER] = '" & Me.Part_Number & "'") > 0 Then 'IF ASSEMBLED
        Set rs1 = db.OpenRecordset("APPS_Q_SIF_NEW_ASSEMBLED_PART_V", dbOpenSnapshot)
        rs1.FindLast "[NIFCO_PART_NUMBER] = '" & Me.Part_Number & "'"
        quoteNumber = Left(rs1!ENG_QUOTE_LOG_NUM, 4)
        sellingPrice = rs1!PIECE_PRICE
        quotedCost = rs1!INTERNAL_PART_COST
    End If
    
    db.Execute "INSERT INTO tblPartQuoteInfo(quoteNumber,quotedCost) VALUES (" & Nz(quoteNumber, 0) & "," & Nz(quotedCost, 0) & ")"
    rsPartInfo.Edit
    rsPartInfo!quoteInfoId = db.OpenRecordset("SELECT @@identity")(0).Value
    rsPartInfo!sellingPrice = sellingPrice
    rsPartInfo.Update
End If

'---Check for Model Code
If Nz(rsPartInfo!programId, 0) = 0 Then
    Dim modelCode As Long
    modelCode = 0
    If Nz(Me.Model_Code, "") <> "" Then modelCode = CLng(Nz(DLookup("ID", "tblPrograms", "modelCode = '" & Me.Model_Code & "'"), 0))
    rsPartInfo.Edit
    rsPartInfo!programId = modelCode
    rsPartInfo.Update
End If

'---Check Molded + Assembly Info---
Select Case Me.ReqTypeID
    Case 9 'new assembly - also add assembly info
        partType = 2
        
        If Nz(rsPartInfo!assemblyInfoId, 0) = 0 Then
            db.Execute "INSERT INTO tblPartAssemblyInfo(assemblyAnnealing) VALUES(0)"
            rsPartInfo.Edit
            rsPartInfo!assemblyInfoId = db.OpenRecordset("SELECT @@identity")(0).Value
            rsPartInfo.Update
        End If
    Case 10 'new molded
        partType = 1
        
        If Nz(rsPartInfo!moldInfoId, 0) = 0 Then
            db.Execute "INSERT INTO tblPartMoldingInfo(toolNumber) VALUES ('" & Me.Part_Number & "T')"
            rsPartInfo.Edit
            rsPartInfo!moldInfoId = db.OpenRecordset("SELECT @@identity")(0).Value
            rsPartInfo.Update
        End If
    Case 11 'new purchased
        partType = 3
End Select

If Nz(rsPartInfo!partType, 0) = 0 Then
    rsPartInfo.Edit
    rsPartInfo!partType = partType
    rsPartInfo.Update
End If

'---Check for tblPartInfo Record on Every Related Part Number---
'-------------------------------------------------------------------------------------------
Set rsRelatedParts = db.OpenRecordset("SELECT * FROM qryRelatedParts WHERE primaryPN = '" & Me.Part_Number & "'")

Do While Not rsRelatedParts.EOF
    If DCount("recordId", "tblPartInfo", "partNumber = '" & rsRelatedParts!relatedPN & "'") = 0 Then db.Execute "INSERT INTO tblPartInfo(partNumber) VALUES ('" & rsRelatedParts!relatedPN & "')"
    rsRelatedParts.MoveNext
Loop

'---Check for partMeetingInfo record---
If DCount("recordId", "tblPartMeetingInfo", "meetingId = " & meetingId) = 0 Then db.Execute "INSERT INTO tblPartMeetingInfo(meetingId) VALUES (" & meetingId & ")"

'---CLEANUP---
On Error Resume Next
rsRelatedParts.CLOSE: Set rsRelatedParts = Nothing
rsPartInfo.CLOSE: Set rsPartInfo = Nothing
rs1.CLOSE: Set rsPartInfo = Nothing
Set db = Nothing
On Error GoTo Err_Handler

DoCmd.OpenForm "frmCrossFunctionalKO", , , "meetId = " & meetingId

'allow edits while open, or always for managers
If (Nz(Me.Completed_Date) = "") Or (userData("Level") = "Manager") Then
    Form_frmCrossFunctionalKO.allowEdits = True
    Form_frmCrossFunctionalKO.copyPI.Enabled = True
Else
    Form_frmCrossFunctionalKO.allowEdits = False
    Form_frmCrossFunctionalKO.copyPI.Enabled = False
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub checker1Sign_Click()
On Error GoTo Err_Handler

If MsgBox("This will remove your signature and all signatures ahead of yours. Are you sure?", vbYesNo, "Woah, nellie") = vbNo Then
    Me.ActiveControl = Me.ActiveControl.OldValue
    Exit Sub
End If

If Me.checker2Sign = True Then 'if checker2 already signed, then remove their signature
    Me.checker2Sign = False
    Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, "checker2Sign", True, False)
    Me.checker2Sign.Locked = True
    Me.checker2SignDate = Null
    Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, "checker2SignDate", Me.checker2SignDate.OldValue, "")
    Me.checker2SignDate.Locked = True
    Me.Check_In_Prog.SetFocus
End If

Me.checker1Sign.Locked = True 'remove checker1 signature
Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, "checker1Sign", True, False)
Me.checker1SignDate = Null
Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, "checker1SignDate", Me.checker1SignDate.OldValue, "")
Me.checker1SignDate.Locked = True

Me.Check_In_Prog.SetFocus
Me.Check_In_Prog = "In Check"
Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, "Check_In_Prog", Me.Check_In_Prog.OldValue, "In Check")
Me.Check_In_Prog.Locked = False
Me.DRSactionButton.Enabled = True
Me.checker1Sign.Visible = False
Me.Repaint

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub checker2Sign_Click()
On Error GoTo Err_Handler

If MsgBox("This will remove your signature. Are you sure?", vbYesNo, "Be careful.") = vbNo Then
    Me.ActiveControl = Me.ActiveControl.OldValue
    Exit Sub
End If

Me.checker2Sign.Locked = True
Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, "checker2Sign", True, False)
Me.checker2SignDate = Null
Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, "checker2SignDate", Me.checker2SignDate.OldValue, "")
Me.checker2SignDate.Locked = True
Me.Check_In_Prog.SetFocus

Me.Check_In_Prog = "In Approval"
Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, "Check_In_Prog", Me.Check_In_Prog.OldValue, "In Approval")
Me.Check_In_Prog.Locked = False
Me.DRSactionButton.Enabled = True

Me.checker2Sign.Visible = False
Me.Repaint
    
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Check_In_Prog_AfterUpdate()
On Error GoTo Err_Handler
'can only freely change between not started and in progress
If ((Me.ActiveControl.OldValue = "Not Started" Or Me.ActiveControl.OldValue = "In Progress") And (Me.ActiveControl = "Not Started" Or Me.ActiveControl = "In Progress")) Then
    Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
Else
    Call snackBox("error", "No can do", "This dropdown will automatically update when using workingDB to sign.", "frmDRSdashboard")
    Me.ActiveControl = Me.ActiveControl.OldValue
    Exit Sub
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub closeDRS_Click()
On Error GoTo Err_Handler

If Len(Me.ECO) > 1 Then
    If IsNull(DLookup("APPROVAL_REQUEST_DATE", "ENG_ENG_ENGINEERING_CHANGES", "Change_Notice = '" & Me.ECO & "'")) Then MsgBox "Hey, don't forget to submit that ECO!", vbInformation, "Just a reminder..."
End If

If Nz(Me.Completed_Date) = "" Then
    MsgBox "You need to enter a due date", vbCritical, "You've been denied"
    Exit Sub
End If

If MsgBox("By entering a completed date, you are closing this DRS. You cannot undo this.", vbOKCancel, "Warning") = vbOK Then
        If Me.Dirty = True Then Me.Dirty = False
        Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.Completed_Date.name, Form_frmDRSdashboard.Completed_Date, Me.Completed_Date)
        
        If DCount("recordId", "tblPartMeetings", "partNum = '" & Me.Part_Number & "' AND meetingType = 1") > 0 Then
            Call scanSteps(Me.Part_Number, "frmCrossFunctionalKO_save") 'now that CFMM are finalized, run check
        End If
        
        DoCmd.CLOSE acForm, "frmDRSDashboard"
        If CurrentProject.AllForms("frmDRSworkTracker").IsLoaded Then Form_frmDRSworkTracker.Requery
End If

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

If Me.Completed_Date < Me.checker2SignDate Then
    Me.Completed_Date = Null
    Call snackBox("error", "Nope", "You must enter a date that is >= to the last approval date", Me.name)
    Exit Sub
End If

Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Delay_Reason_AfterUpdate()
On Error GoTo Err_Handler
Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub drawingChecksheet_Click()
On Error GoTo Err_Handler

Dim controlNum, partType, DRStype, drawingType, designResponsible
controlNum = Me.Control_Number
partType = Nz(Me.DRS_Location) 'DRS_Location is used to store part type
DRStype = Me.ReqTypeID

DoCmd.OpenForm "frmDesignChecksheet", , , "controlNumber = " & controlNum

Dim Assignee As Boolean, Checker As Boolean
Assignee = False
Checker = False
Select Case Environ("username")
    Case Me.Assignee
        Assignee = True
    Case Me.Checker1, Me.Checker2
        Checker = True
End Select

Form_frmDesignChecksheet.btnAssigneeJudgement.Enabled = Assignee
Form_frmDesignChecksheet.btnCheckerJudgement.Enabled = Checker

If IsNull(Me.Completed_Date) = False Then
    Form_frmDesignChecksheet.btnAssigneeJudgement.Enabled = False
    Form_frmDesignChecksheet.btnCheckerJudgement.Enabled = False
    Form_frmDesignChecksheet.Comments.Locked = True
    Form_frmDesignChecksheet.cmbPartType.Locked = True
End If

If DRStype > 3 Then
    drawingType = 1
    Form_frmDesignChecksheet.txtDrawingType = "Internal Drawing"
Else
    drawingType = 2
    Form_frmDesignChecksheet.txtDrawingType = "Customer Drawing"
End If

Select Case Me.DESIGN_RESPONSIBILITY
    Case 1 'Nifco America
        designResponsible = 2
        Form_frmDesignChecksheet.txtDesignResponsible = "NAM Responsible"
    Case 2, 6 'Nifco Japan OR Vtech (NJP)
        designResponsible = 1
        Form_frmDesignChecksheet.txtDesignResponsible = "NJP Responsible"
    Case Else 'Customer only, or customer with our support
        designResponsible = 3
        Form_frmDesignChecksheet.txtDesignResponsible = "Customer Responsible"
End Select

Select Case True
    Case partType = 2 Or DRStype = 9 'Assembled part
        partType = 2
        Form_frmDesignChecksheet.cmbPartType = "Assembled Part"
        Form_frmDesignChecksheet.cmbPartType.Locked = True
    Case partType = 1 Or DRStype = 10 'Molded part
        partType = 1
        Form_frmDesignChecksheet.cmbPartType = "Molded Part"
        Form_frmDesignChecksheet.cmbPartType.Locked = True
    Case partType = 3 Or DRStype = 11 'Purchased part
        partType = 3
        Form_frmDesignChecksheet.cmbPartType = "Purchased Part"
        Form_frmDesignChecksheet.cmbPartType.Locked = True
    Case Else 'partType has not been defined
        Form_frmDesignChecksheet.cmbPartType = "Please Select"
        Form_frmDesignChecksheet.cmbPartType.Locked = False
        Exit Sub
End Select


If DCount("recordId", "tblDesignChecksheet", "controlNumber = " & controlNum) = 0 Then
    Dim db As Database
    Set db = CurrentDb()
    Dim rs1 As Recordset, rsChecksheet As Recordset
    Set rs1 = db.OpenRecordset("SELECT * from tblDesignChecksheetDefaults WHERE drawingType LIKE '*" & drawingType & "*' AND designResponsible LIKE '*" & designResponsible & "*' AND partType LIKE '*" & partType & "*' ORDER BY indexOrder Asc")
    Set rsChecksheet = db.OpenRecordset("tblDesignChecksheet")

    Do While Not rs1.EOF
        rsChecksheet.addNew

        rsChecksheet!controlNumber = controlNum
        rsChecksheet!reviewItem = rs1!recordId

        rsChecksheet.Update
        rs1.MoveNext
    Loop

    rs1.CLOSE
    Set rs1 = Nothing
    Set db = Nothing

    Call registerDRSUpdates("tblDesignChecksheet", controlNum, "Drawing Checksheet", "", "Created", "Created Checksheet", "frmDesignChecksheet")
    Form_frmDesignChecksheet.Requery
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub drsHistory_Click()
On Error GoTo Err_Handler
DoCmd.OpenForm "frmHistory"
Form_frmHistory.RecordSource = "qryDRSupdateTracking"
Form_frmHistory.dataTag0.ControlSource = "dataTag0"
Form_frmHistory.previousData.ControlSource = "previous"
Form_frmHistory.newData.ControlSource = "new"
Form_frmHistory.filter = "tableRecordId = " & Me.Control_Number
Form_frmHistory.FilterOn = True
Form_frmHistory.OrderBy = "updatedDate Desc"
Form_frmHistory.OrderByOn = True
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Due_Date_AfterUpdate()
On Error GoTo Err_Handler
Call registerDRSUpdates("tblDRS", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub ECO_AfterUpdate()
On Error GoTo Err_Handler
Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function setPartPicture()
On Error GoTo Err_Handler
Me.partPicture.Visible = False
If Len(Me.Part_Number) > 4 Then
    Dim partPic, partPicDir
    partPicDir = "\\data\mdbdata\WorkingDB\_docs\Part_Pictures\"
    partPic = Dir((partPicDir & Me.Part_Number & "*"))
    If Len(partPic) > 0 Then
        Me.partPicture.Picture = partPicDir & partPic
        Me.partPicture.Visible = True
    End If
End If
Exit Function
Err_Handler:
    Call handleError(Me.name, "setPartPicture", Err.DESCRIPTION, Err.number)
End Function

Private Sub emailOtherMembers_Click()
On Error GoTo Err_Handler

Dim strTo As String, strCC As String
Dim strSubject As String

Dim cntrlNum, partNum, requestType, dueDate
Dim firstN

'subject line variables
    cntrlNum = Me.Control_Number
    partNum = Me.Part_Number
    requestType = Me.Request_Type
    dueDate = Me.Due_Date

Dim assEm, chk1Em, chk2Em
If Len(Me.Assignee) > 1 Then assEm = getEmail(Me.Assignee) & ";"
If Len(Me.Checker1) > 1 Then chk1Em = getEmail(Me.Checker1) & ";"
If Len(Me.Checker2) > 1 Then chk2Em = getEmail(Me.Checker2) & ";"

Select Case Environ("username")
    Case Me.Assignee
        strTo = chk1Em & chk2Em
    Case Me.Checker2
        strTo = assEm & chk1Em
    Case Me.Checker1
        strTo = assEm & chk2Em
    Case Else
        strTo = assEm & chk1Em & chk2Em
End Select
    
strSubject = cntrlNum & " DRS for " & partNum & " " & requestType & " Due " & dueDate

Call wdbEmail(strTo, "", strSubject, "")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Dim errorTag As Integer
errorTag = 0

Call setTheme(Me)
errorTag = 1

Dim db As Database
Set db = CurrentDb()
errorTag = 2

If (DCount("Control_Number", "tblDRStrackerExtras", "Control_Number = " & TempVars!controlNumber) = 0) Then db.Execute "INSERT INTO tblDRStrackerExtras(Control_Number,Check_In_Prog) VALUES (" & TempVars!controlNumber & ",'Not Started')"
errorTag = 3

Me.Requery
Me.Form.filter = "[Control_Number] = " & TempVars!controlNumber
Me.Form.FilterOn = True
Me.Form.refresh
errorTag = 4

Call setPartPicture
errorTag = 5

Dim closeBool As Boolean, rs1 As Recordset, newStatus As String, chk2 As Boolean, chk1 As Boolean, ass As Boolean, adjustedBool As Boolean, pendingBool As Boolean, closedBool As Boolean
closeBool = False
errorTag = 6

'---show adjusted due date if there is one---
adjustedBool = Not IsNull(Me.Adjusted_Due_Date)
Me.Label85.Visible = adjustedBool
Me.Adjusted_Due_Date.Visible = adjustedBool
errorTag = 7

'---add default tasks if they aren't already in there---
If IsNull(DLookup("[Task_ID]", "[tblTaskTracker]", "[Control_Number] = " & Me.Control_Number)) Then
    Set rs1 = db.OpenRecordset("SELECT * FROM tblTaskTracker WHERE [Control_Number] = 1 AND [Title] = '" & Me.Request_Type & "'", dbOpenSnapshot)
    Do While Not rs1.EOF
        db.Execute "Insert INTO tblTaskTracker(Task, Control_Number, user) VALUES ('" & rs1![Task] & "'," & Me.Control_Number & ",'" & Environ("username") & "');"
        rs1.MoveNext
    Loop
    rs1.CLOSE
    Set rs1 = Nothing
End If
Me.sfrmDRSDashboard.Form.filter = "[Control_Number] = " & Me.Control_Number & " AND [Close_Date] is null"
Me.sfrmDRSDashboard.Form.FilterOn = True
errorTag = 8

'---if still pending, unlock comments and due date---
'---if approved, check for checker2 and allow closing if there isn't one---
pendingBool = Me.Approval_Status = "Pending"
Me.Comments.Locked = Not pendingBool
Me.Due_Date.Locked = Not pendingBool
If IsNull(Me.Checker2) And Not pendingBool Then closeBool = True
errorTag = 9

'---calculate Check Status to account for any errors---
Me.assigneeSign.Locked = True
Me.assigneeSign.Visible = False
Me.checker1Sign.Locked = True
Me.checker1Sign.Visible = False
Me.checker2Sign.Locked = True
Me.checker2Sign.Visible = False
Me.Check_In_Prog.Locked = True
errorTag = 10

If IsNull(Me.Check_In_Prog) Then Me.Check_In_Prog = "Not Started"
chk2 = Me.checker2Sign
chk1 = Me.checker1Sign
ass = Me.assigneeSign
newStatus = Me.Check_In_Prog
errorTag = 11

If ass = False Then 'if assignee hasn't signed, then make sure it's Not Started or In Progress
    If (newStatus <> "Not Started" And newStatus <> "In Progress") Then newStatus = "Not Started"
Else
    Select Case True 'check all signatures and apply correct status
        Case (chk1 And chk2) Or (IsNull(Me.Checker1) And chk2)
            newStatus = "Approved"
            closeBool = True
        Case (chk1 And chk2 = False) Or (IsNull(Me.Checker1) And chk2 = False)
            newStatus = "In Approval"
        Case (chk1 = False And chk2 = False)
            newStatus = "In Check"
    End Select
End If
errorTag = 12

If newStatus <> Me.Check_In_Prog Then
    Me.Check_In_Prog = newStatus
    Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, "Check_In_Prog", Me.Check_In_Prog.OldValue, newStatus)
End If
errorTag = 13

Me.DRSactionButton.Enabled = False
Me.nudgeAssignee.Visible = False
Me.nudgeChecker.Visible = False
Me.nudgeApprover.Visible = False
'Set Action Button availability and caption
Select Case Environ("username")
    Case Me.Assignee
        Me.DRSactionButton.Caption = "Submit WO"
        If Me.assigneeSign = True Then
            Me.assigneeSign.Locked = False
            Me.assigneeSign.Visible = True
        End If
        If Me.Check_In_Prog = "In Progress" Or Me.Check_In_Prog = "Not Started" Then
            Me.Check_In_Prog.Locked = False
            Me.DRSactionButton.Enabled = True
        End If
        Me.nudgeChecker.Visible = True
        Me.nudgeApprover.Visible = True
    Case Me.Checker1
        Me.DRSactionButton.Caption = "Sign as Checker"
        If Me.checker1Sign = True Then
            Me.checker1Sign.Locked = False
            Me.checker1Sign.Visible = True
        End If
        If Me.Check_In_Prog = "In Check" Then Me.DRSactionButton.Enabled = True
        Me.nudgeAssignee.Visible = True
        Me.nudgeApprover.Visible = True
    Case Me.Checker2
        Me.DRSactionButton.Caption = "Sign as Approver"
        If Me.checker2Sign = True Then
            Me.checker2Sign.Locked = False
            Me.checker2Sign.Visible = True
        End If
        If Me.Check_In_Prog = "In Approval" Then Me.DRSactionButton.Enabled = True
        Me.nudgeAssignee.Visible = True
        Me.nudgeChecker.Visible = True
    Case Else
        'this person is not part of the WO
        closeBool = False
        Me.DRSactionButton.Caption = "Not Applicable"
End Select
errorTag = 14

If IsNull(Me.Checker1) Then
    Me.Label159.Visible = True
    Me.nudgeChecker.Visible = False
End If
If Me.Dirty Then Me.Dirty = False

Call setProgressBar
errorTag = 15

'---if WO is already closed, hide things that shouldn't be seen---
Me.frmTimeTrackSummary.Form.AllowAdditions = True
If IsNull(Me.Completed_Date) = False Then
    Me.nudgeAssignee.Visible = False 'can't nudge
    Me.nudgeChecker.Visible = False 'can't nudge
    Me.nudgeApprover.Visible = False 'can't nudge
    Me.btnEditChkFold.Enabled = False 'can't edit check folder
    Me.assigneeSign.Visible = False 'can't unsign
    Me.checker1Sign.Visible = False 'can't unsign
    Me.checker2Sign.Visible = False 'can't unsign
    Me.btnCancel.Enabled = False 'can't cancel
    Me.requestExt.Visible = False 'can't request extension
    Me.DRSactionButton.Caption = "Not Applicable"
    Me.DRSactionButton.Enabled = False
    Me.Completed_Date.Visible = True
    Me.Completed_Date.Locked = True
    closeBool = False
    Me.frmTimeTrackSummary.Form.AllowAdditions = False 'can't add time
End If
errorTag = 16

Me.cfmm.Visible = False
Me.drawingChecksheet.Visible = False
'---For certain WO types, show CFMM or checksheet buttons---
Select Case Me.Request_Type
    Case "New Molded Part", "New Assembly", "New Purchased Part", "Customer Submission", "Customer Re-submission", "Nifco India Support - Internal"
        Me.drawingChecksheet.Visible = True
End Select
errorTag = 17

Select Case Me.Request_Type
    Case "New Molded Part", "New Assembly", "New Purchased Part", "New Replacement Tool"
        Me.cfmm.Visible = True
End Select
errorTag = 18

Set db = Nothing

'---if WO is able to be closed, show the stuff!---
Me.Completed_Date.Locked = Not closeBool
Me.approvalDate.Visible = closeBool
Me.Completed_Date.Visible = closeBool
Me.closeDRS.Visible = closeBool
Me.closeBack.Visible = closeBool
Me.Box198.Visible = closeBool

Me.lblRequestNumber.Visible = Nz(Me.Request_Number, "") <> ""
errorTag = 19

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number, CStr(errorTag))
End Sub

Public Function setProgressBar()
On Error GoTo Err_Handler
Dim percent
Dim pColor
percent = progressPercent(Me.[Control_Number])
If percent < 0.1 Then
    Me.progressBar.Width = 1
Else
    Me.progressBar.Width = percent * 5460
End If

Select Case True
    Case percent < 0.25
        pColor = rgb(210, 110, 90)
    Case percent >= 0.25 And percent < 0.5
        pColor = rgb(225, 170, 70)
    Case percent >= 0.5 And percent < 0.75
        pColor = rgb(200, 210, 100)
    Case percent >= 0.75
        pColor = rgb(125, 215, 100)
End Select
Me.progressBar.BackColor = pColor
Me.progressBar.BorderColor = pColor
Me.progressButton.BorderColor = pColor
Exit Function
Err_Handler:
    Call handleError(Me.name, "setProgressBar", Err.DESCRIPTION, Err.number)
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler

If CurrentProject.AllForms("frmDesignChecksheet").IsLoaded Then DoCmd.CLOSE acForm, "frmDesignChecksheet"
If CurrentProject.AllForms("frmCrossFunctionalKO").IsLoaded Then DoCmd.CLOSE acForm, "frmCrossFunctionalKO"
If CurrentProject.AllForms("frmDRSworkTracker").IsLoaded Or Form_DASHBOARD.appContainer.SourceObject = "frmDRSworkTracker" Then Form_frmDRSworkTracker.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Unload", Err.DESCRIPTION, Err.number)
End Sub

Private Sub fullDRS_Click()
On Error GoTo Err_Handler

If Nz(DLookup("[designWOpermissions]", "tblPermissions", "[user] = '" & Environ("username") & "'"), 3) = 1 Then
    DoCmd.OpenForm "frmApproveDRS", , , "[Control_Number] = " & Me.Control_Number
Else
    Call snackBox("error", "Access Denied", "Only managers can open this form", "frmDRSdashboard")
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub initializeDCR_Click()
On Error GoTo Err_Handler

If populateDCR(Me.Part_Number, Me.Request_Type, Nz(Me.ECO, "")) = False Then
    Call snackBox("error", "Can't make DCR", "DCR Already exists, cannot overwrite", "frmDRSdashboard")
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub meetingOtherMembers_Click()
On Error GoTo Err_Handler

Dim obj0App As Object
Dim objAppt As Object
Dim strTo As String
Dim strSubject As String
Dim cntrlNum, partNum, requestType, dueDate

'subject line variables
    cntrlNum = Me.Control_Number
    partNum = Me.Part_Number
    requestType = Me.Request_Type
    dueDate = Me.Due_Date

Dim assEm, chk1Em, chk2Em
If Len(Me.Assignee) > 1 Then assEm = getEmail(Me.Assignee) & ";"
If Len(Me.Checker1) > 1 Then chk1Em = getEmail(Me.Checker1) & ";"
If Len(Me.Checker2) > 1 Then chk2Em = getEmail(Me.Checker2) & ";"

Select Case Environ("username")
    Case Me.Assignee
        strTo = chk1Em & chk2Em
    Case Me.Checker2
        strTo = assEm & chk1Em
    Case Me.Checker1
        strTo = assEm & chk2Em
    Case Else
        strTo = assEm & chk1Em & chk2Em
End Select
    
strSubject = cntrlNum & " DRS for " & partNum & " " & requestType & " Due " & dueDate

Set obj0App = CreateObject("outlook.Application")
Set objAppt = obj0App.CreateItem(1)

With objAppt
    .RequiredAttendees = strTo
    .subject = strSubject
    .ReminderMinutesBeforeStart = 5
    .Meetingstatus = 1
    .responserequested = True
    .display
End With

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function DRSnudge(sendTo As String, actionLong As String, Action As String)
On Error GoTo Err_Handler

Dim body As String, Comments As String
If Len(Me.Comments) > 25 Then
    Comments = Left(Me.Comments, 25) & "..."
Else
    Comments = Me.Comments
End If
body = emailContentGen("You've been nudged...", "Nudge Notification", "You've been nudged by " & getFullName() & " to " & Action, "WO#" & Me.Control_Number & ", " & Me.Request_Type & " for " & Me.Part_Number, "Comments: " & Comments, "Status: " & Me.Check_In_Prog, "Nudged On: " & CStr(Now()), appName:="Design WO", appId:=Me.Control_Number)
If sendNotification(sendTo, 1, 2, "Please " & actionLong, body, "Design WO", Me.Control_Number) = True Then
    Call snackBox("success", "Well done!", sendTo & " nudged", Me.name)
    Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, "Nudge", "From: " & Environ("username"), "To: " & sendTo)
End If

Exit Function
Err_Handler:
    Call handleError(Me.name, "DRSnudge", Err.DESCRIPTION, Err.number)
End Function

Private Sub nudgeApprover_Click()
On Error GoTo Err_Handler

Dim errorMsg As String
errorMsg = ""

If Me.Check_In_Prog <> "In Approval" Then errorMsg = "Looks like this isn't their problem yet - needs to be 'In Approval' to nudge the approver"
If IsNull(Me.Checker2) Then errorMsg = "No checker found"

If errorMsg <> "" Then
    Call snackBox("error", "Nope", errorMsg, "frmDRSdashboard")
    Exit Sub
End If

Select Case MsgBox("Click YES to nudge this person, click NO to generate new email popup", vbYesNoCancel, "What would you like to do?")
    Case vbYes
        Call DRSnudge(Me.Checker2, "approve WO#" & Me.Control_Number & ", " & Me.Request_Type & " for " & Me.Part_Number, "approve this WO:")
        Me.Check_In_Prog.SetFocus
        Me.nudgeApprover.Visible = False
    Case vbNo
        Call GenerateDRSgeneralEmail(Me.Checker2)
End Select

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub nudgeAssignee_Click()
On Error GoTo Err_Handler

Dim errorMsg As String
errorMsg = ""

If IsNull(Me.Assignee) Then errorMsg = "No one found"

If errorMsg <> "" Then
    Call snackBox("error", "Nope", errorMsg, "frmDRSdashboard")
    Exit Sub
End If

Select Case MsgBox("Click YES to nudge this person, click NO to generate new email popup", vbYesNoCancel, "What would you like to do?")
    Case vbYes
        Call DRSnudge(Me.Assignee, "submit WO#" & Me.Control_Number & ", " & Me.Request_Type & " for " & Me.Part_Number, "submit this WO:")
        Me.Check_In_Prog.SetFocus
        Me.nudgeAssignee.Visible = False
    Case vbNo
        Call GenerateDRSgeneralEmail(Me.Assignee)
End Select

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub GenerateDRSgeneralEmail(person As String)
On Error GoTo Err_Handler
Dim dueDate As String

If IsNull(person) Then
    Call snackBox("error", "No one found", "You can't send an email without a person.", "frmDRSdashboard")
    Exit Sub
End If

If IsNull(Me.Adjusted_Due_Date) Then
    dueDate = Me.Due_Date
Else
    dueDate = Me.Adjusted_Due_Date
End If

Dim objEmail As Object

Set objEmail = CreateObject("outlook.Application")
Set objEmail = objEmail.CreateItem(0)

With objEmail
    .To = getEmail(person)
    .subject = Me.Control_Number & " DRS for " & Me.Part_Number & " " & Me.Request_Type & " Due " & dueDate
    .display
End With

Set objEmail = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub nudgeChecker_Click()
On Error GoTo Err_Handler

Dim errorMsg As String
errorMsg = ""

If IsNull(Me.Checker1) Then errorMsg = "No checker found"
If Me.Check_In_Prog <> "In Check" Then errorMsg = "Please submit for check before you nudge them to check!"

If errorMsg <> "" Then
    Call snackBox("error", "Nope", errorMsg, "frmDRSdashboard")
    Exit Sub
End If

Select Case MsgBox("Click YES to nudge this person, click NO to generate new email popup", vbYesNoCancel, "What would you like to do?")
    Case vbYes
        Call DRSnudge(Me.Checker1, "check WO#" & Me.Control_Number & ", " & Me.Request_Type & " for " & Me.Part_Number, "check this WO:")
        Me.Check_In_Prog.SetFocus
        Me.nudgeChecker.Visible = False
    Case vbNo
        Call GenerateDRSgeneralEmail(Me.Checker1)
End Select

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub refresh_Click()
On Error GoTo Err_Handler
Me.frmTimeTrackSummary.Requery
Me.Requery
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub requestExt_Click()
On Error GoTo Err_Handler

Dim errorMsg As String, errorCount, newDate, reason, oldDate, actionLong, Action, body As String
errorMsg = ""

If Me.Checker2 = "" Or IsNull(Me.Checker2) Then errorMsg = "No approver found"

enterDateError:

newDate = InputBox("Input Date - default is today", "New Due Date Request", Date)
If newDate = "" Then errorMsg = "Nothing enterred"

If errorMsg <> "" Then
    Call snackBox("error", "Not yet", errorMsg, "frmDRSdashboard")
    Exit Sub
End If

errorCount = 0

On Error GoTo enterDateError
errorCount = 1
newDate = CDate(newDate)
On Error GoTo Err_Handler

retry:
reason = InputBox("Input Reason", "New Due Date Request", "")
If StrPtr(reason) = 0 Then Exit Sub
If Len(reason) > 30 Then
    MsgBox "Please enter a shorter reason", vbOKOnly, "Sorry"
    GoTo retry
End If

If IsNull(Me.Adjusted_Due_Date) Or Me.Adjusted_Due_Date = "" Then
    oldDate = CStr(Me.Due_Date)
Else
    oldDate = CStr(Me.Adjusted_Due_Date)
End If

actionLong = ""
Action = ""

body = emailContentGen("Can I please...?", "WO Adjust Request", _
                                            "You've been requested by " & getFullName() & " to adjust due date on WO# " & Me.Control_Number & ", " & Me.Request_Type & " for " & Me.Part_Number, _
                                            "Current: " & oldDate & ", Requested: " & CStr(newDate), _
                                            "Reason: " & reason, _
                                            "", _
                                            "", appName:="Design WO", appId:=Me.Control_Number _
                                        )
If sendNotification(Me.Checker2, 6, 2, "Please adjust due date on WO# " & Me.Control_Number & ", " & Me.Request_Type & " for " & Me.Part_Number, body, "Design WO", Me.Control_Number) = True Then
    Call snackBox("success", "Phew. Hopefully they accept!", "Request sent to " & Me.Checker2 & "!", Me.name)
    Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, "Due Date Ext Request", "From: " & Environ("username"), "To: " & Me.Checker2)
End If

Me.Check_In_Prog.SetFocus
Me.requestExt.Visible = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub notes_AfterUpdate()
On Error GoTo Err_Handler
Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl)
If Me.Dirty Then Me.Dirty = False
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub DRSactionButton_Click()
On Error GoTo Err_Handler
Dim xdate As String, errorMsg As String, chkFold As String, newStatus As String, ctrlName As String

errorMsg = ""

If Me.Approval_Status <> "Approved" Then errorMsg = "WO must be approved to add DRS"

chkFold = getCheckFolder
Me.SetFocus

If chkFold = "" Then errorMsg = "Must have check folder location entered"
If FolderExists(chkFold) = False Then errorMsg = "There may be an issue with your check folder - please check it"

'add check for any approver/checker
If Nz(Me.Checker1) = "" And Nz(Me.Checker2) = "" Then errorMsg = "You need at least an approver to submit for check"

Dim db As Database
Set db = CurrentDb()
Dim rsChecksheet As Recordset, chkSheetExists As Boolean
Set rsChecksheet = db.OpenRecordset("SELECT * FROM tblDesignChecksheet WHERE controlNumber = " & Me.Control_Number)

chkSheetExists = rsChecksheet.RecordCount > 0

If chkSheetExists Then
    Do While Not rsChecksheet.EOF
        If rsChecksheet!assigneeJudgement = False Then errorMsg = "Drawing Checksheet is incomplete."
        rsChecksheet.MoveNext
    Loop
End If

rsChecksheet.CLOSE
Set rsChecksheet = Nothing

'only export for assignee
'only export for new molded, new puchased, new assembly
Dim cfmmId As Long, pnList As String, errorCount As Long
Dim rsRelatedParts As Recordset
Dim rsPI As Recordset

'Check for chkSheet
Select Case Me.Request_Type
    Case "New Purchased Part", "New Molded Part", "New Assembly"
        If chkSheetExists = False Then errorMsg = "You need to fill out the Design Checksheet"
End Select

'Check for CFMM requirements
Select Case Me.Request_Type
    Case "New Molded Part", "New Assembly", "New Replacement Tool"
        cfmmId = Nz(DLookup("recordId", "tblPartMeetings", "partNum = '" & Me.Part_Number & "' AND meetingType = 1"), 0)
        If cfmmId = 0 Then
            errorMsg = "You need to fill out your CFMM"
        Else
            'if there is a meeting, check for the right data
            'first, check this part number's class information
            Set rsPI = db.OpenRecordset("SELECT * FROM tblPartInfo WHERE partNumber = '" & Me.Part_Number & "'")
            
            errorCount = 0
            If Nz(rsPI!partClassCode, 0) = 0 Then errorCount = errorCount + 1
            If Nz(rsPI!subClassCode, 0) = 0 Then errorCount = errorCount + 1
            If Nz(rsPI!businessCode, 0) = 0 Then errorCount = errorCount + 1
            If Nz(rsPI!focusAreaCode, 0) = 0 Then errorCount = errorCount + 1
            
            If errorCount > 0 Then pnList = Me.Part_Number
            
            'then, check all related parts
            Set rsRelatedParts = db.OpenRecordset("SELECT * FROM qryRelatedParts WHERE primaryPN = '" & Me.Part_Number & "'")
            
            Do While Not rsRelatedParts.EOF
                Set rsPI = db.OpenRecordset("SELECT * FROM tblPartInfo WHERE partNumber = '" & rsRelatedParts!relatedPN & "'")
                
                errorCount = 0
                
                If rsPI.RecordCount = 0 Then
                    errorCount = errorCount + 1
                Else
                    If Nz(rsPI!partClassCode, 0) = 0 Then errorCount = errorCount + 1
                    If Nz(rsPI!subClassCode, 0) = 0 Then errorCount = errorCount + 1
                    If Nz(rsPI!businessCode, 0) = 0 Then errorCount = errorCount + 1
                    If Nz(rsPI!focusAreaCode, 0) = 0 Then errorCount = errorCount + 1
                End If
                
                If errorCount > 0 Then pnList = pnList & " " & rsRelatedParts!relatedPN
                rsRelatedParts.MoveNext
            Loop
            
            rsPI.CLOSE
            Set rsPI = Nothing
            rsRelatedParts.CLOSE
            Set rsRelatedParts = Nothing
            
            If pnList <> "" Then errorMsg = "Please fix the class codes for these part numbers: " & pnList
        End If
        
End Select

Set db = Nothing

If errorMsg <> "" Then
    MsgBox errorMsg, vbInformation, "Can't Submit Yet"
    Exit Sub
End If

errorCount = 0
enterDateError:

If (errorCount = 1) Then MsgBox "That's not a date, silly. Let's try that again.", vbQuestion, "Ouch"

xdate = InputBox("Input Date - default is today", "Input Date to Sign DRS", Date)
If xdate = "" Then Exit Sub
 
On Error GoTo enterDateError
errorCount = 1
xdate = CDate(xdate)
On Error GoTo Err_Handler

Me.DRSactionButton.Enabled = False

Select Case Environ("username")
    Case Me.Assignee
        ctrlName = "assigneeSign"
        newStatus = "In Check"
        If IsNull(Me.Checker1) Then newStatus = "In Approval"
    Case Me.Checker1
        If Me.assigneeSign = False Then errorMsg = "Assignee must sign via WorkingDB first!"
        ctrlName = "checker1Sign"
        newStatus = "In Approval"
        If cfmmId <> 0 Then MsgBox "Please make sure to check Class Code on the CFMM", vbInformation, "Please double check"
    Case Me.Checker2
        ctrlName = "checker2Sign"
        If Me.assigneeSign = False Then errorMsg = "Assignee must sign via WorkingDB first!"
        If Not (Me.checker1Sign Or IsNull(Me.Checker1)) Then errorMsg = "Checker must sign via WorkingDB before approver can sign"
        newStatus = "Approved"
        If cfmmId <> 0 Then MsgBox "Please make sure to check Class Code on the CFMM", vbInformation, "Please double check"
End Select

If errorMsg <> "" Then
    MsgBox errorMsg, vbCritical, "Can't Submit Yet"
    Exit Sub
End If

Me.Controls(ctrlName) = True
Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, "ctrlName", Me.Controls(ctrlName).OldValue, Me.Controls(ctrlName))
Me.Controls(ctrlName).Visible = True
Me.Controls(ctrlName & "Date") = xdate
Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, ctrlName & "Date", Me.Controls(ctrlName & "Date").OldValue, Me.Controls(ctrlName & "Date"))
Me.Controls(ctrlName).Locked = False

    
Me.Dirty = False

'---Export PDFs to Check Folder
Dim Y As String, z As String, tempFold As String, fso
Set fso = CreateObject("Scripting.FileSystemObject")
Y = chkFold & "1. " & Me.Control_Number & " DRS for " & Me.Part_Number & " " & Me.Request_Type & " Due " & Format(Me.Due_Date, "mmddyyyy") & ".pdf"

tempFold = getTempFold
If FolderExists(tempFold) = False Then MkDir (tempFold)

'Export DRS
z = tempFold & Me.Control_Number & "TEMP.pdf"
DoCmd.OpenReport "rptDesignRequest", acViewPreview, , "[Control_Number]=" & Me.Control_Number, acHidden
DoCmd.OutputTo acOutputReport, "rptDesignRequest", acFormatPDF, z, False
DoCmd.CLOSE acReport, "rptDesignRequest"

Call fso.CopyFile(z, Y)
Call fso.deleteFile(z)

MsgBox "DRS Signed", vbOKOnly, "File Added to Check Folder"

'Export Checksheet IF checksheet is filled out
'checksheet gets exported IF: new assembly, new molded, new puchased, customer submission (sometimes), customer resubmission (sometimes)
If chkSheetExists Then
    Y = chkFold & "2. " & Me.Control_Number & " Design Checksheet.pdf"
    
    z = tempFold & Me.Control_Number & "TEMPdesignChecksheet.pdf"
    DoCmd.OpenReport "rptDesignChecksheet", acViewPreview, , "[controlNumber]=" & Me.Control_Number, acHidden
    DoCmd.OutputTo acOutputReport, "rptDesignChecksheet", acFormatPDF, z, False
    DoCmd.CLOSE acReport, "rptDesignChecksheet"
    
    Call fso.CopyFile(z, Y)
    Call fso.deleteFile(z)

    MsgBox "Checksheet Signed", vbOKOnly, "File Added to Check Folder"
End If

'Export CFMM IF CFMM is filled out
If cfmmId <> 0 Then
    Y = chkFold & "6. " & Me.Part_Number & " CFMM.pdf"
    
    z = tempFold & Me.Control_Number & "TEMPcfmm.pdf"
    DoCmd.OpenReport "rptCrossFunctionalKO", acViewPreview, , "[meetingId]=" & cfmmId, acHidden
    DoCmd.OutputTo acOutputReport, "rptCrossFunctionalKO", acFormatPDF, z, False
    DoCmd.CLOSE acReport, "rptCrossFunctionalKO"
    
    Call fso.CopyFile(z, Y)
    Call fso.deleteFile(z)

    MsgBox "CFMM Exported", vbOKOnly, "File Added to Check Folder"
End If

If Environ("username") = Me.Assignee Then Call approvalEmail

Me.Check_In_Prog = newStatus
Me.Check_In_Prog.Locked = True
Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, "Check_In_Prog", Me.Check_In_Prog.OldValue, newStatus)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function approvalEmail()
On Error GoTo Err_Handler

Dim strTo As String, strCC As String
Dim strSubject As String

Dim cntrlNum, partNum, requestType, dueDate
Dim firstN
Dim chkFold As String

'subject line variables
    cntrlNum = Me.Control_Number
    partNum = Me.Part_Number
    requestType = Me.Request_Type
    dueDate = Me.Due_Date

'recipient / body variables
    firstN = DLookup("[firstName]", "tblPermissions", "[user] = '" & Me.Checker1 & "'")

If IsNull(Me.Checker1) Or Me.Checker1 = "" Then
        strTo = getEmail(Me.Checker2)
Else
        strTo = getEmail(Me.Checker1)
        strCC = getEmail(Me.Checker2)
End If
    chkFold = "file:///" & getCheckFolder
    chkFold = Replace(chkFold, "''", "''''")
    
    strSubject = cntrlNum & " DRS for " & partNum & " " & requestType & " Due " & dueDate

Dim body As String
Dim subTitle As String, detail1 As String, detail2 As String, detail3 As String
subTitle = "This Work Order is ready for check. " & getFullName() & " has submitted this WO"
detail1 = "Request Type: " & requestType
detail2 = "Part Number: " & partNum
detail3 = "Control Number: " & cntrlNum
body = generateHTML("WO Submission", subTitle, "Check Folder", detail1, detail2, detail3, chkFold, True)
If wdbEmail(strTo, strCC, strSubject, body) = True Then Call snackBox("success", "Success!", "WO Submitted", Me.name)

Exit Function
Err_Handler:
    Call handleError(Me.name, "approvalEmail", Err.DESCRIPTION, Err.number)
End Function

Private Sub searchPN_Click()
On Error GoTo Err_Handler

Form_DASHBOARD.partNumberSearch = Me.Part_Number
Form_DASHBOARD.filterbyPN_Click
Form_DASHBOARD.SetFocus

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub WOcatiaMacros_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmCatiaMacros"
Form_frmCatiaMacros.tabWO.Visible = True
Form_frmCatiaMacros.tabWO.SetFocus

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
