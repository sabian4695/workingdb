Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub clear_Click()
On Error GoTo Err_Handler
Me.NAMsrchBox.SetFocus
Me.NAMsrchBox = ""
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltOrg_AfterUpdate()
On Error GoTo Err_Handler

applyFilter

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Function applyFilter(Optional frozenOnly As Boolean = False)

Dim partNum, filt
partNum = Me.NAMsrchBox
If partNum = "" Then Exit Function

filt = "[ITEM_NUMBER] = '" & partNum & "'"
If frozenOnly Then filt = filt & " AND COST_TYPE = 'Frozen'"
If Me.fltOrg <> "ALL" Then filt = filt & " AND Org = '" & Me.fltOrg & "'"

DoCmd.applyFilter , filt

End Function

Private Sub FormHeader_Click()
On Error GoTo Err_Handler

applyFilter

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub NAMsrch_Click()
On Error GoTo Err_Handler

applyFilter

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub srchFrozen_Click()
On Error GoTo Err_Handler

applyFilter (True)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
