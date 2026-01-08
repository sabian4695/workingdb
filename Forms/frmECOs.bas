Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub clear_Click()
On Error GoTo Err_Handler
Me.ECOsrch.SetFocus
Me.ECOsrch = ""
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub PersonSrch_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmECObyPerson"
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub revItemSrch_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmECOpartHistory"
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Sub srch_Click()
On Error GoTo Err_Handler

Me.filter = "[CHANGE_NOTICE] = '" & UCase(Me.ECOsrch) & "'"
Me.FilterOn = True
Me.Repaint

If Me.RecordsetClone.RecordCount = 0 Then
        Me.filter = "[Change_Notice] = 'CNL10000'"
        Me.FilterOn = True
        MsgBox "There are no records" & vbCrLf & "that match this filter.", vbInformation + vbOKOnly, "No records returned"
        Exit Sub
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
