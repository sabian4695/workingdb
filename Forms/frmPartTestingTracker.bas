Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function filterIt(controlName As String)
On Error GoTo Err_Handler

Me(controlName).SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Function
Err_Handler:
    Call handleError(Me.name, "filterIt", Err.DESCRIPTION, Err.number)
End Function

Private Sub actualend_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTesting", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub actualStart_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTesting", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub addNew_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False

If restrict(Environ("username"), "New Model Quality") Then Exit Sub 'only NMQ

If IsNull(Me.fltPartNumber) Then
    MsgBox "Please select a part number in the filter dropdown first!", vbInformation, "Fix this first"
    Exit Sub
End If

Dim db As Database
Set db = CurrentDb()

db.Execute "INSERT INTO tblPartTesting(projectId,partNumber) VALUES (" & Me.fltPartNumber.column(1) & ",'" & Me.fltPartNumber & "');"
TempVars.Add "testId", db.OpenRecordset("SELECT @@identity")(0).Value
Call registerPartUpdates("tblPartTesting", TempVars!testId, "Test", "", "Created", Me.fltPartNumber, Me.name, Me.fltPartNumber.column(1))
Me.Requery

Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub allHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory", acNormal, , "[tableName] = 'tblPartTesting'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnOpenWO_Click()
On Error GoTo Err_Handler

If Nz(Me.workOrderNumber) = "" Then Exit Sub

If CurrentProject.AllForms("frmLabWOs").IsLoaded = False Then
    DoCmd.OpenForm "frmLabWOs"
End If

Form_frmLabWOs.txtSearch = Me.workOrderNumber

Form_frmLabWOs.Form.filter = "WONumber = " & Replace(Form_frmLabWOs.txtSearch, "N", "")
Form_frmLabWOs.Form.FilterOn = True

DoCmd.CLOSE acForm, "frmLabWO_Search"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub customerSpec_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTesting", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub duration_AfterUpdate()
On Error GoTo Err_Handler

Me.plannedEnd = Me.plannedStart + Me.duration
Call registerPartUpdates("tblPartTesting", Me.recordId, Me.plannedEnd.name, Me.plannedEnd.OldValue, Me.plannedEnd, Me.partNumber, Me.name)

Me.duration = Null

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub export_Click()
On Error GoTo Err_Handler

Dim FileName As String, sqlString As String
FileName = "H:\Testing_" & nowString & ".xlsx"
sqlString = "SELECT tblPartTesting.partNumber, tblDropDownsSP.testType, tblPartTesting.customerSpec, tblPartTesting.workOrderNumber, tblPartTesting.plannedStart, " & _
                    "tblPartTesting.plannedEnd, tblPartTesting.actualStart, tblPartTesting.actualEnd, tblPartTesting.pass, tblPartTesting.notes " & _
                    "FROM tblPartTesting LEFT JOIN tblDropDownsSP ON tblPartTesting.testType = tblDropDownsSP.recordid  where " & Me.Form.filter

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

If Me.showPassToggle.Value = False Then filt = "testStatus <> 3"

If Me.fltPartNumber <> "" Then
    If filt <> "" Then filt = filt & " AND "
    filt = filt & "partNumber = '" & Me.fltPartNumber & "'"
    Me.fltUser = Null
    Me.fltModel = Null
    GoTo filtNow
End If

If Me.fltUser <> "" Then
    If filt <> "" Then filt = filt & " AND "
    filt = filt & "partNumber IN (SELECT partNumber FROM tblPartTeam WHERE person = '" & Me.fltUser & "')"
End If

If Me.fltModel <> "" Then
    If filt <> "" Then filt = filt & " AND "
    filt = filt & "partNumber IN (SELECT partNumber FROM tblPartInfo WHERE programId = " & Me.fltModel & ")"
End If

filtNow:
Me.filter = filt
Me.FilterOn = filt <> ""

Exit Function
Err_Handler:
    Call handleError(Me.name, "applyTheFilters", Err.DESCRIPTION, Err.number)
End Function

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
'NMQ manager can edit any
'NMQ must be on CF team
If (DCount("partNumber", "tblPartTeam", "person = '" & Environ("username") & "'") > 0) And (Not restrict(Environ("username"), "New Model Quality")) Then Exit Sub 'on CF team and NMQ
If Not restrict(Environ("username"), "New Model Quality", "Manager") Then Exit Sub 'if NMQ manager

MsgBox "You must be NMQ and on CF team or NMQ Manager to edit", vbCritical, "Nope"
Me.Undo

End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim allowEdits As Boolean
allowEdits = (Not restrict(Environ("username"), "New Model Quality"))

Me.addNew.Enabled = allowEdits

Me.OrderBy = "plannedStart"
Me.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub notes_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTesting", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub notes1_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTesting", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openDetails_Click()
On Error GoTo Err_Handler

openPartProject (Me.partNumber)

DoCmd.CLOSE acForm, "frmPartTestingTracker"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub plannedEnd_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTesting", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub plannedStart_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTesting", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)

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

Private Sub remove_Click()
On Error GoTo Err_Handler

If (restrict(Environ("username"), "New Model Quality") = True) Then
    MsgBox "Only NMQ can remove this", vbCritical, "Denied"
    Exit Sub
End If

If MsgBox("This will delete this entire test item, not just the notes. Are you sure you want to delete this?", vbYesNo, "Please confirm") = vbYes Then
    Call registerPartUpdates("tblPartTesting", Me.recordId, "Test", Me.testType, "Deleted", Me.partNumber, Me.name, Me.projectId)
    dbExecute ("DELETE FROM tblPartTesting WHERE [recordId] = " & Me.recordId)
    Me.Requery
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub testClassification_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTesting", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub testFiles_Click()
On Error GoTo Err_Handler

If IsNull(Me.testType) Then
    MsgBox "Please enter test type before adding attachments", vbCritical, "Hold on"
    Exit Sub
End If

DoCmd.OpenForm "frmPartAttachments", , , "testId = " & Me.recordId
Form_frmPartAttachments.TpartNumber = Me.partNumber
Form_frmPartAttachments.TtestId = Me.recordId
Form_frmPartAttachments.TprojectId = Me.projectId
Form_frmPartAttachments.itemName.Caption = Me.testType.column(1)
Form_frmPartAttachments.secondaryType = 22

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub testHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory", acNormal, , "[tableName] = 'tblPartTesting' AND [partNumber] = '" & Me.partNumber & "' AND [tableRecordId] = " & Me.recordId

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub testStatus_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTesting", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub testType_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTesting", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub workOrderNumber_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartTesting", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
