Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub byPN_Click()
On Error GoTo Err_Handler

Dim partNum
partNum = Me.NAMsrchBox
If partNum = "" Then Exit Sub

DoCmd.applyFilter , "[SEGMENT1] = '" & partNum & "'"
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub byTool_Click()
On Error GoTo Err_Handler

Dim partNum
partNum = Me.NAMsrchBox
If partNum = "" Then Exit Sub

DoCmd.applyFilter , "[RESOURCE_CODE] = '" & partNum & "'"
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub clear_Click()
On Error GoTo Err_Handler
Me.NAMsrchBox.SetFocus
Me.NAMsrchBox = ""
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
