Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Function checkInputs() As String

checkInputs = ""

Select Case True
    Case Nz(Me.stepName) = ""
        checkInputs = "Title"
    Case Nz(Me.responsible) = ""
        checkInputs = "Responsible"
    Case Nz(Me.stepNotes) = ""
        checkInputs = "Step Notes"
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

dbExecute ("UPDATE tblCPC_Steps SET indexOrder = indexOrder + 1 WHERE projectId = " & Me.projectId & " AND indexOrder >= " & Me.indexOrder)

Call registerCPCUpdates("tblCPC_Steps", Me.ID, "Created", "", "Created", Me.projectId, Me.stepName)

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
    If Nz(Me.ID) <> "" Then DoCmd.RunCommand acCmdDeleteRecord
    DoCmd.SetWarnings True
End If

Form_sfrmCPC_Dashboard.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Unload", Err.DESCRIPTION, Err.number)
End Sub
