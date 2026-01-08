Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Function checkInputs() As String

checkInputs = ""

Select Case True
    Case Nz(Me.stepType) = ""
        checkInputs = "Title"
    Case Nz(Me.responsible) = ""
        checkInputs = "Responsible"
    Case Nz(Me.DESCRIPTION) = ""
        checkInputs = "Description"
End Select

End Function

Private Sub btnSave_Click()
On Error GoTo Err_Handler

Dim val As String
val = checkInputs

If val <> "" Then
    MsgBox "Must Enter " & val, vbInformation, "Fix it"
    Exit Sub
End If
If Me.Dirty = True Then Me.Dirty = False

'update indeces
Dim indexVal As Long
indexVal = Me.indexVal

dbExecute ("UPDATE tblPartSteps SET indexOrder = indexOrder + 1 WHERE partGateId = " & Form_sfrmPartDashboard.partGateId & " AND indexOrder > " & indexVal)

Me.indexOrder = indexVal + 1

Call registerPartUpdates("tblPartSteps", Me.recordId, "Created", "", "Created", Me.partNumber, Me.stepType, Me.partProjectId)

DoCmd.CLOSE

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.openedBy = Environ("username")
Me.lastUpdatedBy = Environ("username")
Me.responsible = userData("Dept")
Me.responsible.DefaultValue = userData("Dept")

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler

If checkInputs() <> "" Then
    If MsgBox("Are you sure?" & vbNewLine & "Your current record will be deleted.", vbYesNo, "Please confirm") <> vbYes Then
        Cancel = True
        Exit Sub
    End If

    DoCmd.SetWarnings False
    If Nz(Me.recordId) <> "" Then DoCmd.RunCommand acCmdDeleteRecord
    DoCmd.SetWarnings True
End If

If CurrentProject.AllForms("frmPartDashboard").IsLoaded Then Form_frmPartDashboard.partDash_refresh_Click

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Unload", Err.DESCRIPTION, Err.number)
End Sub
