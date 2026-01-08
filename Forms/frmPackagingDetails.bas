Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

If Not Me.Dirty Then Exit Sub
If Not restrict(Environ("username"), "Packaging") Then Exit Sub 'if packaging engineer, then OK

MsgBox "You must be Packaging to edit", vbCritical, "Nope"
Me.Undo

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_BeforeUpdate", Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim packagingE As Boolean
packagingE = Not restrict(Environ("username"), "Packaging")

Me.packRank.Locked = Not packagingE
Me.packagingTest.Locked = Not packagingE
Me.fitTrialStatus.Locked = Not packagingE
Me.partsExpected.Locked = Not packagingE
Me.totesAllocated.Locked = Not packagingE
Me.totesNeeded.Locked = Not packagingE
Me.totesStatus.Locked = Not packagingE
Me.Notes.Locked = Not packagingE

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler

Form_frmPackagingTracker.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Unload", Err.DESCRIPTION, Err.number)
End Sub

Private Sub openDetails_Click()
On Error GoTo Err_Handler

openPartProject (Me.partNumber)
DoCmd.CLOSE acForm, "frmPartPackagingDetails"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function afterUpdate_tblPartPackagingInfo()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartPackagingInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)

If Me.Dirty Then Me.Dirty = False

Exit Function
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Function

Private Sub packHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory", acNormal, , "[tableName] = 'tblPartPackagingInfo' AND [tableRecordId] = " & Me.recordId

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub save_Click()
On Error GoTo Err_Handler

DoCmd.CLOSE acForm, "frmPartPackagingDetails"

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
