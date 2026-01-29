Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function applyTheFilters()
On Error GoTo Err_Handler
Dim filt

filt = ""

If Me.showClosedToggle.Value Then
        filt = "[closeDate] is not null"
    Else
        filt = "[closeDate] is null"
End If
    

If Me.fltPartNumber <> "" Then
    If filt <> "" Then filt = filt & " AND "
    filt = filt & "partNumber = '" & Me.fltPartNumber & "'"
    Me.fltInCharge = Null
    Me.fltModel = Null
    Me.filtOpenedBy = Null
    GoTo filtNow
End If

If Me.fltInCharge <> "" Then
    If filt <> "" Then filt = filt & " AND "
    filt = filt & "inCharge = '" & Me.fltInCharge & "'"
End If

If Me.filtOpenedBy <> "" Then
    If filt <> "" Then filt = filt & " AND "
    filt = filt & "issueOpenedBy = '" & Me.filtOpenedBy & "'"
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

Private Sub btnImportPhoto_Click()
On Error GoTo Err_Handler

Dim fd As FileDialog
Dim FileName As String
    
Set fd = Application.FileDialog(msoFileDialogOpen)
With fd
    .Filters.clear
    .Filters.Add "Images", "*.png; *.gif; *.jpg; *.jpeg"
    .InitialFileName = "C:\Users\Public"
End With
    
fd.Show
On Error GoTo errorCatch
FileName = fd.SelectedItems(1)

Dim general, partNum

general = "\\data\mdbdata\WorkingDB\_docs\Part_Issues_Docs\"
partNum = Me.partNumber

If FolderExists(general & partNum) = True Then
    GoTo pathMade
Else
    MkDir (general & partNum)
    GoTo pathMade
End If

pathMade:
    Dim fso, FilePath
    Set fso = CreateObject("Scripting.FileSystemObject")
    FilePath = general & partNum & "\" & partNum & "_issue_" & Me.recordId & ".png"
    Call fso.CopyFile(FileName, FilePath)
    Me.issuePic.Picture = FilePath

Me.export.SetFocus
Me.btnImportPhoto.Visible = False

errorCatch:
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub closeDate_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartIssues", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub closeIssue_Click()
On Error GoTo Err_Handler

If IsNull(Me.issueCountermeasure) Then
    MsgBox "Please enter a countermeasure", vbInformation, "Not yet you won't"
    Exit Sub
End If

If MsgBox("Are you sure?", vbYesNo, "Please confirm") <> vbYes Then Exit Sub

Me.closeDate = Now()
Me.issueStatus = 3
Call registerPartUpdates("tblPartIssues", Me.recordId, Me.closeDate, Me.closeDate.OldValue, Me.closeDate, Me.partNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub details_Click()
On Error GoTo Err_Handler

If (CurrentProject.AllForms("frmPartIssueDetails").IsLoaded = True) Then
    DoCmd.CLOSE acForm, "frmPartIssueDetails"
End If
DoCmd.OpenForm "frmPartIssueDetails", , , "recordId = " & Me.recordId

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub dueDate_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartIssues", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub export_Click()
On Error GoTo Err_Handler

Dim FileName As String, sqlString As String, filt As String
FileName = "H:\Issues_" & nowString & ".xlsx"
filt = " WHERE " & Me.Form.filter
If Me.FilterOn = False Then filt = ""
sqlString = "SELECT partNumber, tblDropDownsSP.issueType, tblDropDownsSP_1.issueSource, foundBy, foundDate, issueDescription, " & _
                        "tblDropDownsSP_2.issuePriority, issueOpenedDate, issueOpenedBy, inCharge, tblDropDownsSP_3.issueStatus, closeDate, issueCountermeasure " & _
                        "FROM (((tblPartIssues INNER JOIN tblDropDownsSP ON tblPartIssues.issueType = tblDropDownsSP.recordid) INNER JOIN tblDropDownsSP AS tblDropDownsSP_1 ON tblPartIssues.issueSource = tblDropDownsSP_1.recordid) " & _
                        "INNER JOIN tblDropDownsSP AS tblDropDownsSP_2 ON tblPartIssues.issuePriority = tblDropDownsSP_2.recordid) INNER JOIN tblDropDownsSP AS tblDropDownsSP_3 ON tblPartIssues.issueStatus = tblDropDownsSP_3.recordid " & filt

Call exportSQL(sqlString, FileName)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub filtOpenedBy_AfterUpdate()
On Error GoTo Err_Handler
applyTheFilters
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltInCharge_AfterUpdate()
On Error GoTo Err_Handler
applyTheFilters
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltModel_AfterUpdate()
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
applyTheFilters
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Current()
On Error GoTo Err_Handler


Dim notClosed As Boolean
notClosed = IsNull(Me.closeDate)

Me.closeIssue.Visible = notClosed
Me.closeDate.Visible = Not notClosed
Me.Label130.Visible = Not notClosed
Me.btnImportPhoto.Visible = notClosed
Me.remove.Visible = notClosed

Me.issueCountermeasure.Locked = Not notClosed
Me.issueType.Locked = Not notClosed
Me.issueSource.Locked = Not notClosed
Me.inCharge.Locked = Not notClosed
Me.issueStatus.Locked = Not notClosed
Me.issuePriority.Locked = Not notClosed
Me.issueDescription.Locked = Not notClosed
Me.foundBy.Locked = Not notClosed
Me.foundDate.Locked = Not notClosed
Me.closeDate.Locked = Not notClosed

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub foundBy_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartIssues", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub foundDate_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartIssues", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Image90_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmUserProfile", , , "user = '" & Me.inCharge & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub imgFoundBy_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmUserProfile", , , "user = '" & Me.foundBy & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub imgInCharge_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmUserProfile", , , "user = '" & Me.inCharge & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub imgOpenedBy_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmUserProfile", , , "user = '" & Me.issueOpenedBy & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub inCharge_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartIssues", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber)

If (Me.inCharge <> Me.inCharge.OldValue) And (Me.inCharge <> Environ("username")) Then 'if this is an actual change, shoot the new assignee the email
    Dim emailString As String
    emailString = emailContentGen("New Open Issue Assignment", "You've been assigned an Open Issue", Me.issueDescription, "Due " & Me.dueDate, _
        Me.issuePriority.column(1) & " Priority", "Source: " & Me.issueSource.column(1) & "; Type: " & Me.issueType.column(1), "Part Number: " & Me.partNumber)
    Call sendNotification(Me.inCharge, 4, 2, "You've been assigned an Open Issue", emailString, "Issue", Me.recordId)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub issueCountermeasure_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartIssues", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub issueDescription_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartIssues", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub issueHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory", acNormal, , "[tableName] = 'tblPartIssues' AND [tableRecordId] = " & Me.recordId

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub issuePriority_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartIssues", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub issueSource_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartIssues", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub issueStatus_AfterUpdate()
On Error GoTo Err_Handler

If Me.issueStatus = 3 Then
    Me.Undo
    Call snackBox("error", "Please, no", "Use the 'Close Issue' button to close an issue!", "DASHBOARD")
    Exit Sub
End If

Call registerPartUpdates("tblPartIssues", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub issueType_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartIssues", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function validate() As Boolean

validate = True

Select Case True
    Case IsNull(Me.partNumber)
        MsgBox "Please enter a part number.", vbOKOnly, "Warning"
        validate = False
    Case IsNull(Me.issueType)
        MsgBox "Please enter an issue type", vbOKOnly, "Warning"
        validate = False
    Case IsNull(Me.issueSource)
        MsgBox "Please enter an issue source.", vbOKOnly, "Warning"
        validate = False
    Case IsNull(Me.issuePriority)
        MsgBox "Please enter an issue priority.", vbOKOnly, "Warning"
        validate = False
    Case IsNull(Me.foundBy)
        MsgBox "Please enter who found this issue. ", vbOKOnly, "Warning"
        validate = False
    Case IsNull(Me.foundDate)
        MsgBox "Please enter when this issue was found. ", vbOKOnly, "Warning"
        validate = False
    Case IsNull(Me.inCharge)
        MsgBox "Please enter who is in charge. ", vbOKOnly, "Warning"
        validate = False
    Case IsNull(Me.issueDescription)
        MsgBox "Please enter an issue description", vbOKOnly, "Warning"
        validate = False
End Select

End Function

Private Sub lblDue_Click()
On Error GoTo Err_Handler

Me.rowDueDate.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblID_Click()
On Error GoTo Err_Handler

Me.recordId.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblInCharge_Click()
On Error GoTo Err_Handler

Me.rowInCharge.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblOpened_Click()
On Error GoTo Err_Handler

Me.issueOpenedDate.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblOpenedBy_Click()
On Error GoTo Err_Handler

Me.issueOpenedBy.SetFocus
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

Private Sub lblPriority_Click()
On Error GoTo Err_Handler

Me.rowIssuePriority.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblSource_Click()
On Error GoTo Err_Handler

Me.rowIssueSource.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblStatus_Click()
On Error GoTo Err_Handler

Me.rowIssueStatus.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblType_Click()
On Error GoTo Err_Handler

Me.rowIssueType.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub newPartIssue_Click()
On Error GoTo Err_Handler
    
If IsNull(Me.fltPartNumber) Then
    MsgBox "Please select a part number in the filter dropdown first!", vbInformation, "Fix this first"
    Exit Sub
End If

Me.export.SetFocus
Me.newPartIssue.Visible = False
If Me.Dirty Then Me.Dirty = False

Dim db As Database
Set db = CurrentDb()

db.Execute "INSERT INTO tblPartIssues(issueOpenedBy,inCharge,foundBy,partNumber, issueStatus, issuePriority,foundDate) VALUES " & _
                                    "('" & Environ("username") & "','" & Environ("username") & "','" & Environ("username") & "','" & Me.fltPartNumber & "',1,2,Date());"
TempVars.Add "issueId", db.OpenRecordset("SELECT @@identity")(0).Value
Call registerPartUpdates("tblPartIssues", TempVars!issueId, "Issue Creation", "", "Created", Me.fltPartNumber, Me.name)
Me.Requery

Me.filter = "recordId = " & TempVars!issueId
Me.FilterOn = True

Me.save.Visible = True

Set db = Nothing

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

Public Sub refresh_Click()
On Error GoTo Err_Handler
Me.Requery
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub remove_Click()
On Error GoTo Err_Handler

If MsgBox("Are you sure you want to delete this?", vbYesNo, "Please confirm") = vbYes Then
    Call registerPartUpdates("tblPartIssues", Me.recordId, "Issue", Me.issueType.column(1), "Deleted", Me.partNumber, Me.name)
    dbExecute ("DELETE FROM tblPartIssues WHERE [recordId] = " & Me.recordId)
    Me.Requery
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub save_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False
If validate = False Then Exit Sub

Dim body As String
body = generateHTML(Me.partNumber & " Issue Created", _
    Me.issueDescription, Me.issueType.column(1) & " " & Me.issueSource.column(1) & " Issue Added", _
    Me.issueType.column(1) & " " & Me.issueSource.column(1), "Part Number: " & Me.partNumber, _
    "Created By: " & getFullName, appName:="Issue", appId:=Me.recordId)
    
    
Dim partTeam As String
partTeam = grabPartTeam(partNumber, onlyEngineers:=True, includeMe:=True)
If partTeam <> "" Then Call sendNotification(partTeam, 9, 2, "Open Issue Created for " & Me.partNumber, body, "Issue", Me.recordId, True)

applyTheFilters
Me.export.SetFocus
Me.save.Visible = False
Me.newPartIssue.Visible = True

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
