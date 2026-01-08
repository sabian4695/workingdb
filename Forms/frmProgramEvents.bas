Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnProgramDetails_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmProgramReview"

Dim Program
Program = Me.modelCode

Form_frmProgramReview.txtFilterInput.Value = Program
Form_frmProgramReview.filterByProgram_Click

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub correlatedGate_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblProgramEvents", Me.programId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub dataSubmitted_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblProgramEvents", Me.programId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub dataSubmittedDate_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblProgramEvents", Me.programId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub eventTitle_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblProgramEvents", Me.programId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub eventDate_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblProgramEvents", Me.programId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function applyFilter()
On Error GoTo Err_Handler
Dim filt
filt = ""

'filters are exclusive, except for "show past" toggle button
If Me.ActiveControl.name = "filtOEM" Then Me.filtModel = Null
If Me.ActiveControl.name = "filtModel" Then Me.filtOEM = Null

If Not Me.showClosedToggle Then filt = filt & "eventDate > Date()"

If Nz(Me.filtOEM) <> "" Then
    If filt <> "" Then filt = filt & " AND "
    filt = filt & "OEM = '" & Me.filtOEM & "'"
    Me.filtModel = ""
    GoTo filtNow
End If

If Nz(Me.filtModel) <> "" Then
    If filt <> "" Then filt = filt & " AND "
    filt = filt & "modelCode LIKE '*" & Me.filtModel & "*'"
    Me.filtOEM = ""
    GoTo filtNow
End If

filtNow:
Me.filter = filt
Me.FilterOn = filt <> ""

Exit Function
Err_Handler:
    Call handleError(Me.name, "applyTheFilters", Err.DESCRIPTION, Err.number)
End Function

Private Sub eventType_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblProgramEvents", Me.programId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub filtModel_AfterUpdate()
On Error GoTo Err_Handler

Call applyFilter

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub filtOEM_AfterUpdate()
On Error GoTo Err_Handler

Call applyFilter

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.filter = "eventDate > Date()"
Me.FilterOn = True

Dim lockElsePE As Boolean, lockElseNMQ As Boolean

lockElsePE = restrict(Environ("username"), "Project", "Supervisor", True)
lockElseNMQ = restrict(Environ("username"), "New Model Quality", "Supervisor", True)

Me.eventDate.Locked = lockElsePE
Me.eventTitle.Locked = lockElsePE
Me.correlatedGate.Locked = lockElsePE
Me.eventType.Locked = lockElsePE
Me.dataSubmitted.Locked = lockElseNMQ
Me.dataSubmittedDate.Locked = lockElseNMQ

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub showClosedToggle_Click()
On Error GoTo Err_Handler

Call applyFilter

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
