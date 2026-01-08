Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub allHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory", acNormal, , "[tableName] = 'tblPartTrials'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub engineerStartup_Click()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTrials", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name, Nz(Me.projectId))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub export_Click()
On Error GoTo Err_Handler

Dim FileName As String, sqlString As String, filt As String
FileName = "H:\Trials_" & nowString & ".xlsx"
filt = " WHERE " & Replace(Me.Form.filter, "partNumber", "tblPartTrials.partNumber")
filt = Replace(filt, "trialStatus", "tblPartTrials.trialStatus")
filt = Replace(filt, "frmPartTrialTracker", "tblPartTrials")
If Me.FilterOn = False Then filt = ""
sqlString = "SELECT * FROM qryPartTrialTrackerExport " & filt
                    
Call exportSQL(sqlString, FileName)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltModel_AfterUpdate()
On Error GoTo Err_Handler
Me.fltPartNumber = Null
applyTheFilters
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function applyTheFilters()
On Error GoTo Err_Handler
Dim filt

filt = ""

If Not Me.showClosedToggle.Value Then
    filt = "[trialStatus] <> 3"
End If

If Not Me.showCancelledToggle.Value Then
    If filt <> "" Then filt = filt & " AND "
    filt = filt & "[trialStatus] <> 4"
End If

If Me.fltPartNumber <> "" Then
    If filt <> "" Then filt = filt & " AND "
    filt = "partNumber = '" & Me.fltPartNumber & "'"
    Me.fltUser = Null
    Me.fltModel = Null
    GoTo filtNow
End If

If Me.fltUser <> "" Then
    If filt <> "" Then filt = filt & " AND "
    filt = filt & "(partNumber IN (SELECT partNumber FROM tblPartTeam WHERE person = '" & Me.fltUser & "') OR (processingEngineer = '" & Environ("username") & "'))"
End If

If Me.fltModel <> "" Then
    If filt <> "" Then filt = filt & " AND "
    filt = filt & "partNumber IN (SELECT partNumber FROM tblPartInfo WHERE programId = " & Me.fltModel & ")"
End If

If Me.fltOrg <> "" Then
    If filt <> "" Then filt = filt & " AND "
    filt = filt & "partNumber IN (SELECT partNumber FROM tblPartInfo WHERE developingLocation = '" & Me.fltOrg & "')"
End If

filtNow:
Me.filter = filt
Me.FilterOn = filt <> ""

Exit Function
Err_Handler:
    Call handleError(Me.name, "applyTheFilters", Err.DESCRIPTION, Err.number)
End Function

Private Sub fltOrg_AfterUpdate()
On Error GoTo Err_Handler
applyTheFilters
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltPartNumber_AfterUpdate()
On Error GoTo Err_Handler
applyTheFilters
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltUser_AfterUpdate()
On Error GoTo Err_Handler
Me.fltPartNumber = Null
applyTheFilters
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

If Not Me.Dirty Then Exit Sub

If Not restrict(Environ("username"), "Processing") Then Exit Sub 'if processing engineer, then OK

MsgBox "You must be a Processing Engineer", vbCritical, "Nope"
Me.Undo

End Sub

Function validate()

validate = False

Dim errorMsg As String
errorMsg = ""

If IsNull(Me.trialDate) Then errorMsg = "Trial Date"
If IsNull(Me.press) Then errorMsg = "Press"

If Me.trialResult = 2 Then 'NG
    If Nz(Me.failureType, "") = "" Then errorMsg = "Failure Type"
    If Nz(Me.failureMode, "") = "" Then errorMsg = "Failure Mode"
End If

If errorMsg <> "" Then
    MsgBox "Please fill out " & errorMsg, vbInformation, "Please fix"
    Exit Function
End If

validate = True

End Function

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim allowEdits As Boolean
allowEdits = (Not restrict(Environ("username"), "Processing")) Or (Not restrict(Environ("username"), "Project"))

'PE and Proc E can create/edit. ONLY Proc E can edit Status
Me.newTrial.Enabled = allowEdits
Me.publishTrials.Enabled = allowEdits
Me.Ttrial.Locked = Not allowEdits
Me.trialDate.Locked = Not allowEdits
Me.press.Locked = Not allowEdits
Me.trialStatus.Locked = restrict(Environ("username"), "Processing")
Me.trialResult.Locked = Not allowEdits
Me.trialReason.Locked = Not allowEdits

Me.OrderBy = "trialDate Desc"
Me.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblCurrentUnit_Click()
On Error GoTo Err_Handler

Me.currentUnit.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblEngineerStartup_Click()
On Error GoTo Err_Handler

Me.engineerStartup.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblNotes_Click()
On Error GoTo Err_Handler

Me.Notes.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblPartStatus_Click()
On Error GoTo Err_Handler

Me.partStatus.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblPN_Click()
On Error GoTo Err_Handler

Me.partNumber.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblPress_Click()
On Error GoTo Err_Handler

Me.press.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblProcEngingeer_Click()
On Error GoTo Err_Handler

Me.processingEngineer.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblReason_Click()
On Error GoTo Err_Handler

Me.trialReason.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblResult_Click()
On Error GoTo Err_Handler

Me.trialResult.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblStatus_Click()
On Error GoTo Err_Handler

Me.trialStatus.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblTN_Click()
On Error GoTo Err_Handler

Me.toolNumber.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblTrialDate_Click()
On Error GoTo Err_Handler

Me.trialDate.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lbltTrial_Click()
On Error GoTo Err_Handler

Me.Ttrial.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub newTrial_Click()
On Error GoTo Err_Handler

If restrict(Environ("username"), "Processing") And restrict(Environ("username"), "Project") Then Exit Sub 'only Processing

If IsNull(Me.fltPartNumber) Then
    MsgBox "Please select a part number in the filter dropdown first!", vbInformation, "Fix this first"
    Exit Sub
End If

Me.fltModel.SetFocus
If Me.Dirty Then Me.Dirty = False

Dim db As Database
Set db = CurrentDb()

'find partInfoId
Dim partInfoId As String, pNum As String
partInfoId = Nz(DLookup("recordId", "tblPartInfo", "partNumber = '" & Me.fltPartNumber & "'"), "")

If Len(Me.fltPartNumber) > 5 And Right(Me.fltPartNumber, 1) = "T" Then
    pNum = Left(Me.fltPartNumber, 5)
Else
    pNum = Me.fltPartNumber
End If

'add part info if it doesn't exist
If partInfoId = "" Then
    Form_DASHBOARD.partNumberSearch = pNum
    TempVars.Add "partNumber", pNum
    
    'run Dash search
    Form_DASHBOARD.filterbyPN_Click
    
    If Form_DASHBOARD.lblErrors.Visible = True And Form_DASHBOARD.lblErrors.Caption = "Part not found in Oracle" Then
        MsgBox "This part number must show up in Oracle", vbInformation, "Sorry."
        GoTo exit_handler
    End If
    
    Dim X, defaultT As String
    If Right(Me.fltPartNumber, 1) <> "T" Then
        defaultT = Me.fltPartNumber & "T"
    Else
        defaultT = Me.fltPartNumber
    End If
    X = InputBox("Please confirm the tool number", "Tool Number", defaultT)
    If X = "" Or X = vbCancel Then Exit Sub
    
    'first, does this tool exist? if so, find it and put it in.
    Dim moldInfoId
    moldInfoId = Nz(DLookup("recordId", "tblPartMoldingInfo", "toolNumber = '" & X & "'"), "")
    If moldInfoId = "" Then
        db.Execute "INSERT INTO tblPartMoldingInfo(toolNumber) VALUES ('" & X & "')"
        moldInfoId = db.OpenRecordset("SELECT @@identity")(0).Value
    End If
    
    Dim rsPartInfo As Recordset
    Set rsPartInfo = db.OpenRecordset("tblPartInfo")
    
    rsPartInfo.addNew
    rsPartInfo!partNumber = Me.fltPartNumber
    rsPartInfo!developingLocation = DLookup("permissionsLocation", "tblDropDownsSP", "recordid = " & userData("Org"))
    rsPartInfo!moldInfoId = moldInfoId
    rsPartInfo.Update
    
    partInfoId = db.OpenRecordset("SELECT @@identity")(0).Value
    
    rsPartInfo.CLOSE
    Set rsPartInfo = Nothing
End If

'find current unit
Dim invId, currentUnit As String, rsCat As Recordset
invId = Nz(idNAM(pNum, "NAM"), "")

currentUnit = ""

If invId <> "" Then
    Set rsCat = db.OpenRecordset("SELECT SEGMENT1 FROM INV_MTL_ITEM_CATEGORIES LEFT JOIN APPS_MTL_CATEGORIES_VL ON INV_MTL_ITEM_CATEGORIES.CATEGORY_ID = APPS_MTL_CATEGORIES_VL.CATEGORY_ID " & _
    "GROUP BY INV_MTL_ITEM_CATEGORIES.INVENTORY_ITEM_ID, APPS_MTL_CATEGORIES_VL.SEGMENT1, APPS_MTL_CATEGORIES_VL.STRUCTURE_ID HAVING STRUCTURE_ID = 101 AND [INVENTORY_ITEM_ID] = " & invId, dbOpenSnapshot)
    If rsCat.RecordCount > 0 Then currentUnit = Nz(rsCat!SEGMENT1, "")

    rsCat.CLOSE
    Set rsCat = Nothing
End If

'find part status
Dim rsStatus As Recordset, partStatus As String
Set rsStatus = db.OpenRecordset("SELECT INVENTORY_ITEM_STATUS_CODE FROM APPS_MTL_SYSTEM_ITEMS WHERE SEGMENT1 = '" & pNum & "'", dbOpenSnapshot)

partStatus = ""
If rsStatus.RecordCount > 0 Then partStatus = Nz(rsStatus!INVENTORY_ITEM_STATUS_CODE, "")

rsStatus.CLOSE
Set rsStatus = Nothing

Dim tTrialNumber As Long
tTrialNumber = Nz(DMax("Ttrial", "tblPartTrials", "trialStatus <> 4 AND engineerStartup = true AND partNumber = '" & Me.fltPartNumber & "'"), -1) + 1

Dim runnerWeight As Double
runnerWeight = Nz(DLookup("runnerWeight", "tblPartTrials", "trialStatus <> 4 AND engineerStartup = true AND partNumber = '" & Me.fltPartNumber & "'"), 0)

db.Execute "INSERT INTO tblPartTrials(partNumber,creator,partInfoId,currentUnit,partStatus,Ttrial,processingEngineer,runnerWeight) VALUES ('" & _
    Me.fltPartNumber & "','" & _
    Environ("username") & "'," & _
    partInfoId & ",'" & _
    currentUnit & "','" & _
    partStatus & "'," & _
    tTrialNumber & ",'" & _
    Environ("username") & "'," & _
    runnerWeight & ");"
TempVars.Add "trialId", db.OpenRecordset("SELECT @@identity")(0).Value
Call registerPartUpdates("tblPartTrials", TempVars!trialId, "Trial", "", "Created", Me.fltPartNumber, Me.name, Nz(Me.fltPartNumber.column(1)))
Me.Requery

Set db = Nothing

DoCmd.OpenForm "frmPartTrialDetails", acNormal, , "recordId = " & TempVars!trialId

exit_handler:
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
GoTo exit_handler
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

DoCmd.OpenForm "frmPartTrialDetails", , , "recordId = " & Me.recordId

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

Private Sub publishTrials_Click()
On Error GoTo Err_Handler

'go through all trials for TODAY that are SCHEDULED in this DEV LOCATION
'check that org is filtered, then remove all other filters and use the forms recordset

If Nz(Me.fltOrg) = "" Then
    MsgBox "Please select an Org in the filter area, and try again", vbInformation, "Try Again"
    Exit Sub
End If

Me.fltPartNumber = ""
Me.fltUser = ""
Me.fltModel = ""

Me.filter = "trialDate >= Date() AND trialStatus = 2 AND partNumber IN (SELECT partNumber FROM tblPartInfo WHERE developingLocation = '" & Me.fltOrg & "') "
Me.FilterOn = True

Dim rs1 As Recordset
Set rs1 = Me.RecordsetClone

rs1.Sort = "trialDate"

Set rs1 = rs1.OpenRecordset

Dim dataArray() As Variant, i As Long
ReDim Preserve dataArray(10, 0)
i = 0

dataArray(0, 0) = "Trial<br/>Date"
dataArray(1, 0) = "PN<br/>Tool"
dataArray(2, 0) = "Desc."
dataArray(3, 0) = "Press"
dataArray(4, 0) = "Mat#"
dataArray(5, 0) = "Engineers"
dataArray(6, 0) = "Oracle Job"
dataArray(7, 0) = "Parts<br/>Needed"
dataArray(8, 0) = "Reason"
dataArray(9, 0) = "Model"
dataArray(10, 0) = "Notes"

Dim k As Integer
k = 1

Dim db As Database
Set db = CurrentDb()

Dim rsTeam As Recordset

Do While Not rs1.EOF
    ReDim Preserve dataArray(10, k)
    'find team members
    
    Dim TE, PE, NMQ, Material, rsPerm As Recordset
    Set rsTeam = db.OpenRecordset("select * from tblPartTeam where partNumber = '" & rs1!partNumber & "' AND person is not null")
    TE = ""
    PE = ""
    NMQ = ""
    
    Do While Not rsTeam.EOF
        Set rsPerm = db.OpenRecordset("SELECT * from tblPermissions WHERE user = '" & rsTeam!person & "'")
        If rsPerm!Level = "Engineer" Then
            Select Case rsPerm!dept
                Case "Tooling"
                    TE = rsTeam!person
                Case "Project"
                    PE = rsTeam!person
                Case "New Model Quality"
                    NMQ = rsTeam!person
            End Select
        End If
        
        rsTeam.MoveNext
    Loop
    
    If Nz(rs1!alternateMaterial) <> "" Then
        Material = "ALT: " & rs1!alternateMaterial
    Else
        Material = rs1!materialNumber
        If Nz(rs1!materialNumber1) <> "" Then Material = Material & " + " & rs1!materialNumber1
    End If
    
    Dim qtyNeed
    If Nz(rs1!oracleJob) = "" Then
        qtyNeed = ""
    Else
        qtyNeed = DLookup("START_QUANTITY", "APPS_WIP_DISCRETE_JOBS_V", "WIP_ENTITY_NAME = '" & rs1!oracleJob & "'")
    End If
    
    Dim partDesc As String
    If Nz(rs1!DESCRIPTION, "") = "" Then
        partDesc = findDescription(rs1!partNumber)
    Else
        partDesc = rs1!DESCRIPTION
    End If

    dataArray(0, k) = rs1!trialDate
    dataArray(1, k) = rs1!partNumber & "<br/>" & rs1!toolNumber
    dataArray(2, k) = partDesc
    dataArray(3, k) = rs1!press
    dataArray(4, k) = Material
    dataArray(5, k) = "TE: " & TE & "<br/>PE: " & PE & "<br/>NMQ: " & NMQ
    dataArray(6, k) = Nz(rs1!oracleJob)
    dataArray(7, k) = qtyNeed
    dataArray(8, k) = DLookup("trialReason", "tblDropdownsSP", "recordid = " & rs1!trialReason)
    dataArray(9, k) = DLookup("modelCode", "tblPrograms", "ID = " & Nz(rs1!programId, "0"))
    dataArray(10, k) = Left(rs1!Notes, 300)
    
    k = k + 1
    rs1.MoveNext
Loop

Dim htmlBody As String
htmlBody = trialScheduleEmail("Sample Trial Schedule", dataArray, 10, k - 1)

Call wdbEmail("", "", "Sample Trial Schedule", htmlBody)

Call registerPartUpdates("tblPartTrials", 0, "Publish Trials", "", "Publish Trials", "", Me.name)

On Error Resume Next
rsTeam.CLOSE
Set rsTeam = Nothing
rs1.CLOSE
Set rs1 = Nothing
rsPerm.CLOSE
Set rsPerm = Nothing
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub refresh_Click()
On Error GoTo Err_Handler
Me.Requery
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

Private Sub showCancelledToggle_Click()
On Error GoTo Err_Handler
applyTheFilters
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub showClosedToggle_Click()
On Error GoTo Err_Handler
applyTheFilters
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

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub trialStatus_AfterUpdate()
On Error GoTo Err_Handler

If Me.trialStatus = 3 Then 'complete
    'in order to mark a trial as complete, you must have cycle time / runner w / piece w
    Dim errorMsg
    errorMsg = ""
    
    If Nz(Me.cycleTime, 0) = 0 Then errorMsg = "Cycle Time"
    If Nz(Me.pieceWeight, 0) = 0 Then errorMsg = "Piece Weight"
    If Nz(Me.Notes, "") = "" Then errorMsg = "Notes"
    If IsNull(Me.runnerWeight) Then errorMsg = "Runner Weight"
    
    If errorMsg <> "" Then
        MsgBox "Please fill out " & errorMsg, vbInformation, "Please fix"
        Me.trialStatus = Me.trialStatus.OldValue
        Exit Sub
    End If
End If

Call registerPartUpdates("tblPartTrials", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name, Nz(Me.projectId))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Ttrial_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTrials", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name, Nz(Me.projectId))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
