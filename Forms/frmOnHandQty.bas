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

Private Sub NAMsrch_Click()
On Error GoTo Err_Handler

Dim partNum
partNum = Me.NAMsrchBox
If partNum <> "" Then partNum = idNAM(partNum, "NAM")
If partNum = "" Then Exit Sub

DoCmd.applyFilter , "[INVENTORY_ITEM_ID] = " & partNum
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
